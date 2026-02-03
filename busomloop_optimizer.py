"""
Busomloop Optimizer - NS Trein Vervangend Vervoer (TVV)
=======================================================
Leest het invoer-Excel bestand (Bijlage J) in en genereert:
  1. Busomloop per bus (Transvision-stijl)
  2. Overzicht van ritsamenhang
  3. Berekeningen en optimalisatie-details

Gebruik:
    python busomloop_optimizer.py <invoer.xlsx> [--output <uitvoer.xlsx>]

Keertijden worden automatisch bepaald uit de data (kleinste gap per bustype,
minimum 2 minuten). Handmatig overschrijven kan met:
    --keer-dd 15  --keer-tc 8  --keer-lvb 12  --keer-midi 10  --keer-taxi 5
"""

import argparse
import datetime
import sys
from dataclasses import dataclass, field
from typing import Optional

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class Trip:
    trip_id: str
    bus_nr: int
    service: str          # Sheet / busdienst name
    date_str: str         # e.g. "do 11-06-2026"
    date_label: str       # e.g. "donderdag 11 juni"
    direction: str        # "heen" or "terug"
    bus_type: str         # Dubbeldekker, Touringcar, Taxibus
    snel_stop: str        # snelbus / stopbus
    pattern: str
    multiplicity: int     # Aantal bussen
    origin_code: str
    origin_name: str
    origin_halt: str
    dest_code: str
    dest_name: str
    dest_halt: str
    departure: int        # minutes from midnight
    arrival: int          # minutes from midnight
    stops: list           # list of (station_code, station_name, halt_code, halt_name, time_minutes)
    copy_nr: int = 1      # which copy (1..multiplicity)

    @property
    def duration(self) -> int:
        return self.arrival - self.departure

    def time_str(self, minutes: int) -> str:
        h, m = divmod(minutes % 1440, 60)
        return f"{h:02d}:{m:02d}"

    @property
    def dep_str(self) -> str:
        return self.time_str(self.departure)

    @property
    def arr_str(self) -> str:
        return self.time_str(self.arrival)


@dataclass
class BusRotation:
    bus_id: str
    bus_type: str
    date_str: str
    trips: list = field(default_factory=list)

    @property
    def start_time(self) -> int:
        return self.trips[0].departure if self.trips else 0

    @property
    def end_time(self) -> int:
        return self.trips[-1].arrival if self.trips else 0

    @property
    def total_ride_minutes(self) -> int:
        return sum(t.arrival - t.departure for t in self.trips)

    @property
    def total_idle_minutes(self) -> int:
        idle = 0
        for i in range(1, len(self.trips)):
            gap = self.trips[i].departure - self.trips[i - 1].arrival
            idle += gap
        return idle

    @property
    def total_dienst_minutes(self) -> int:
        return self.end_time - self.start_time if self.trips else 0


@dataclass
class ReserveBus:
    station: str
    count: int
    day: str
    start: int       # minutes from midnight
    end: int         # minutes from midnight
    remark: str = ""


# ---------------------------------------------------------------------------
# Time helpers
# ---------------------------------------------------------------------------

def time_to_minutes(t) -> Optional[int]:
    """Convert datetime.time or decimal fraction to minutes from midnight."""
    if t is None:
        return None
    if isinstance(t, datetime.time):
        return t.hour * 60 + t.minute
    if isinstance(t, (int, float)):
        # Could be decimal day fraction
        total = round(t * 24 * 60)
        return total
    return None


def minutes_to_str(m: int) -> str:
    if m is None:
        return ""
    h, mi = divmod(m % 1440, 60)
    return f"{h:02d}:{mi:02d}"


def minutes_to_time(m: int) -> datetime.time:
    h, mi = divmod(m % 1440, 60)
    return datetime.time(h, mi)


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------

def parse_reserve_buses(wb) -> list:
    """Parse reserve bus info from Voorblad."""
    ws = wb["Voorblad"]
    reserves = []
    in_reserve = False
    for row in ws.iter_rows(min_row=1, max_row=100, max_col=15, values_only=False):
        cell_vals = {c.column: c.value for c in row}
        if cell_vals.get(1) == "Reservebussen":
            in_reserve = True
            continue
        if in_reserve and cell_vals.get(2) == "Station":
            continue  # header row
        if in_reserve and cell_vals.get(2) and cell_vals.get(4):
            station = cell_vals.get(2, "")
            count = int(cell_vals.get(4, 0))
            day = cell_vals.get(5, "")
            start_t = time_to_minutes(cell_vals.get(6))
            end_t = time_to_minutes(cell_vals.get(8))
            remark = cell_vals.get(10, "") or ""
            if start_t is not None and end_t is not None:
                reserves.append(ReserveBus(station, count, day, start_t, end_t, remark))
    return reserves


BUS_TYPE_VALUES = {"dubbeldekker", "touringcar", "taxibus", "lagevloerbus",
                    "gelede bus", "midi bus", "midibus"}


def normalize_bus_type(raw: str) -> str:
    """Normalize bus type names to canonical form matching the pricing categories."""
    low = raw.strip().lower()
    if "dubbeldek" in low:
        return "Dubbeldekker"
    if "touring" in low:
        return "Touringcar"
    if "lagevloer" in low or "gelede" in low:
        return "Lagevloerbus"
    if "midi" in low:
        return "Midi bus"
    if "taxi" in low:
        return "Taxibus"
    # Return with title case if unknown
    return raw.strip() if raw.strip() else "Onbekend"


def parse_direction_block(ws, start_row, max_col=100):
    """
    Parse a 'Busbewegingen in heen/terugrichting' block.
    Returns list of dicts with trip info per bus column.
    """
    bus_numbers = {}
    patterns = {}
    bus_types = {}
    snel_stop = {}
    aantal = {}
    looptijd = {}
    stops = []

    for row in ws.iter_rows(min_row=start_row, max_row=start_row + 80,
                            max_col=max_col, values_only=False):
        cell_map = {c.column: c.value for c in row}
        label = str(cell_map.get(5, "") or "").strip()
        col1 = cell_map.get(1, "")

        data_vals = {col: val for col, val in cell_map.items()
                     if col >= 6 and val is not None}

        if label == "Busnummer":
            bus_numbers = data_vals
        elif label == "Patroon":
            # Detect what this "Patroon" row actually contains
            first_val = next(iter(data_vals.values()), "") if data_vals else ""
            if isinstance(first_val, (int, float)) and not isinstance(first_val, datetime.time):
                # Integer values = these are bus numbers
                bus_numbers = data_vals
            elif str(first_val).strip().lower() in BUS_TYPE_VALUES:
                bus_types = data_vals
            elif isinstance(first_val, datetime.time):
                # Time values = this is actually Looptijd
                looptijd = {col: time_to_minutes(val) for col, val in data_vals.items()}
            else:
                patterns = data_vals
        elif label == "Type bus":
            bus_types = data_vals
        elif "Snelbus" in label or "Stopbus" in label:
            snel_stop = data_vals
        elif label == "Aantal bussen":
            aantal = data_vals
        elif "Looptijd" in label:
            looptijd = {col: time_to_minutes(val) for col, val in data_vals.items()}
        elif cell_map.get(2) and cell_map.get(3):
            # Station row - must have station code in col 2 and name in col 3
            station_code = cell_map.get(2, "")
            station_name = cell_map.get(3, "")
            halt_code = cell_map.get(4, "")
            halt_name = cell_map.get(5, "")
            rijtijd = cell_map.get(1, "")
            times = {}
            for col, val in data_vals.items():
                t = time_to_minutes(val)
                if t is not None:
                    times[col] = t
            stops.append({
                "station_code": str(station_code),
                "station_name": str(station_name),
                "halt_code": str(halt_code) if halt_code else "",
                "halt_name": str(halt_name) if halt_name else "",
                "rijtijd": rijtijd,
                "times": times,
            })
        elif col1 and str(col1).startswith("Uurpatronen"):
            break
        elif col1 and str(col1).startswith("Busbewegingen"):
            break

    return bus_numbers, patterns, bus_types, snel_stop, aantal, looptijd, stops


def parse_sheet(wb, sheet_name) -> list:
    """Parse a single service sheet and return list of Trip objects."""
    ws = wb[sheet_name]
    trips = []

    # Read busdienst and datum
    service = ""
    date_str = ""
    for row in ws.iter_rows(min_row=1, max_row=5, max_col=5, values_only=False):
        cell_map = {c.column: c.value for c in row}
        if cell_map.get(1) == "Busdienst":
            service = str(cell_map.get(2, ""))
        elif cell_map.get(1) == "Datum":
            date_str = str(cell_map.get(2, ""))

    # Find direction blocks
    max_col = ws.max_column or 20
    if max_col > 100:
        max_col = 100

    heen_start = None
    terug_start = None
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=1, max_row=ws.max_row or 100,
                     max_col=5, values_only=False), start=1
    ):
        cell_map = {c.column: c.value for c in row}
        val = cell_map.get(1, "")
        if val and "Busbewegingen in heenrichting" in str(val):
            heen_start = row_idx + 1
        elif val and "Busbewegingen in terugrichting" in str(val):
            terug_start = row_idx + 1

    trip_counter = 0

    for direction, start_row in [("heen", heen_start), ("terug", terug_start)]:
        if start_row is None:
            continue
        bus_numbers, patterns, bus_types_map, snel_stop_map, aantal_map, looptijd_map, stops = \
            parse_direction_block(ws, start_row, max_col)

        if not bus_numbers or not stops:
            continue

        # For each bus column, build a trip
        for col, bus_nr in sorted(bus_numbers.items()):
            # Check this column has times in stops
            col_times = []
            for s in stops:
                t = s["times"].get(col)
                if t is not None:
                    col_times.append((s, t))

            if len(col_times) < 2:
                continue

            origin = col_times[0]
            dest = col_times[-1]
            all_stops = [(s["station_code"], s["station_name"],
                          s["halt_code"], s["halt_name"], t)
                         for s, t in col_times]

            multiplicity = int(aantal_map.get(col, 1))
            bus_type = normalize_bus_type(str(bus_types_map.get(col, "Onbekend")))
            pattern = str(patterns.get(col, ""))
            snel = str(snel_stop_map.get(col, ""))

            # Handle times that cross midnight
            dep_time = origin[1]
            arr_time = dest[1]
            if arr_time < dep_time:
                arr_time += 1440  # next day

            for copy in range(1, multiplicity + 1):
                trip_counter += 1
                trip_id = f"{sheet_name}_{direction}_{bus_nr}_{copy}"

                t = Trip(
                    trip_id=trip_id,
                    bus_nr=bus_nr,
                    service=service,
                    date_str=date_str,
                    date_label=sheet_name,
                    direction=direction,
                    bus_type=bus_type,
                    snel_stop=snel,
                    pattern=pattern,
                    multiplicity=multiplicity,
                    origin_code=origin[0]["station_code"],
                    origin_name=origin[0]["station_name"],
                    origin_halt=origin[0]["halt_name"],
                    dest_code=dest[0]["station_code"],
                    dest_name=dest[0]["station_name"],
                    dest_halt=dest[0]["halt_name"],
                    departure=dep_time,
                    arrival=arr_time,
                    stops=all_stops,
                    copy_nr=copy,
                )
                trips.append(t)

    return trips


def parse_all_sheets(input_file: str):
    """Parse all service sheets from the input Excel."""
    wb = openpyxl.load_workbook(input_file, data_only=True)
    all_trips = []
    reserves = parse_reserve_buses(wb)

    for sheet_name in wb.sheetnames:
        if sheet_name == "Voorblad":
            continue
        try:
            sheet_trips = parse_sheet(wb, sheet_name)
            all_trips.extend(sheet_trips)
        except Exception as e:
            print(f"  Waarschuwing: Fout bij parsen van '{sheet_name}': {e}")

    return all_trips, reserves, wb.sheetnames


# ---------------------------------------------------------------------------
# Optimizer - Greedy best-fit bus chaining
# ---------------------------------------------------------------------------

# Default minimum turnaround times per bus type (minutes)
MIN_TURNAROUND_DEFAULTS = {
    "Dubbeldekker": 15,
    "Touringcar": 8,
    "Lagevloerbus": 12,
    "Midi bus": 10,
    "Taxibus": 5,
}
MIN_TURNAROUND_FALLBACK = 8  # fallback for unknown bus types
MIN_TURNAROUND_FLOOR = 2     # absolute minimum turnaround (minutes)


def detect_turnaround_times(trips: list, within_service_only: bool = False) -> dict:
    """
    Auto-detect minimum turnaround time per bus type from the trip data.

    If within_service_only=True, only considers trips from the same service
    (= same Excel tab). This gives a conservative baseline turnaround since
    it avoids accidental short gaps between unrelated services.

    If within_service_only=False, considers all trips at the same location
    regardless of service (the original behavior).

    Returns dict {bus_type: minutes}, with a floor of MIN_TURNAROUND_FLOOR.
    """
    # Group arrivals and departures
    arrivals = {}
    departures = {}

    for t in trips:
        dest_loc = normalize_location(t.dest_code)
        orig_loc = normalize_location(t.origin_code)
        if within_service_only:
            arr_key = (t.bus_type, t.date_str, dest_loc, t.service)
            dep_key = (t.bus_type, t.date_str, orig_loc, t.service)
        else:
            arr_key = (t.bus_type, t.date_str, dest_loc)
            dep_key = (t.bus_type, t.date_str, orig_loc)
        arrivals.setdefault(arr_key, []).append(t.arrival)
        departures.setdefault(dep_key, []).append(t.departure)

    # For each bus type, find minimum gap between any arrival and subsequent departure
    min_gap_per_type = {}

    for arr_key, arr_times in arrivals.items():
        bus_type = arr_key[0]
        dep_key = arr_key  # same key structure
        if dep_key not in departures:
            continue

        dep_times = sorted(departures[dep_key])
        for arr_t in arr_times:
            for dep_t in dep_times:
                gap = dep_t - arr_t
                if gap >= MIN_TURNAROUND_FLOOR:
                    if bus_type not in min_gap_per_type or gap < min_gap_per_type[bus_type]:
                        min_gap_per_type[bus_type] = gap
                    break

    result = {}
    for bus_type in set(t.bus_type for t in trips):
        if bus_type in min_gap_per_type:
            result[bus_type] = min_gap_per_type[bus_type]
        else:
            result[bus_type] = MIN_TURNAROUND_FALLBACK

    return result


def detect_turnaround_per_service(trips: list) -> dict:
    """
    Detect the minimum turnaround time per service (= per Excel tab).
    Returns dict {service_name: (bus_type, min_gap_minutes)}.
    """
    by_service = {}
    for t in trips:
        by_service.setdefault(t.service, []).append(t)

    result = {}
    for service, svc_trips in by_service.items():
        bus_type = svc_trips[0].bus_type if svc_trips else "Onbekend"
        arrivals = {}
        departures = {}
        for t in svc_trips:
            dest_loc = normalize_location(t.dest_code)
            orig_loc = normalize_location(t.origin_code)
            arrivals.setdefault(dest_loc, []).append(t.arrival)
            departures.setdefault(orig_loc, []).append(t.departure)

        min_gap = None
        for loc, arr_times in arrivals.items():
            dep_times = departures.get(loc, [])
            dep_sorted = sorted(dep_times)
            for arr_t in arr_times:
                for dep_t in dep_sorted:
                    gap = dep_t - arr_t
                    if gap >= MIN_TURNAROUND_FLOOR:
                        if min_gap is None or gap < min_gap:
                            min_gap = gap
                        break

        result[service] = (bus_type, min_gap if min_gap is not None else MIN_TURNAROUND_FALLBACK)

    return result


def normalize_location(code: str) -> str:
    """Normalize station codes for matching (e.g. same city = same location)."""
    code = code.strip().lower()
    # Map variants to canonical form
    mapping = {
        "ah": "arnhem", "ah90": "arnhem", "ah92": "arnhem",
        "ed": "ede", "ed93": "ede",
        "ut": "utrecht", "ut92": "utrecht", "uto": "utrecht_overvecht",
        "db": "driebergen", "db90": "driebergen",
        "klp": "veenendaal_klomp", "klp90": "veenendaal_klomp",
        "rhn": "rhenen", "rhn90": "rhenen",
        "vdn": "veenendaal", "vdnc": "veenendaal_centrum",
        "mrn": "maarn", "mrn90": "maarn", "mrn91": "maarn",
        "amf": "amersfoort", "amf91": "amersfoort",
        "bhv": "bilthoven", "bhv90": "bilthoven",
        "dld": "den_dolder", "dld90": "den_dolder",
        "bkl": "breukelen", "bkl90": "breukelen",
        "htn": "houten", "htn90": "houten",
        "hvs": "houten_vinex", "gdm": "geldermalsen",
        "wd": "woerden", "wd90": "woerden",
    }
    return mapping.get(code, code)


def can_connect(prev_trip: Trip, next_trip: Trip, turnaround_map: dict) -> bool:
    """Check if a bus finishing prev_trip can start next_trip."""
    # Must be same bus type
    if prev_trip.bus_type != next_trip.bus_type:
        return False
    # Must be same date
    if prev_trip.date_str != next_trip.date_str:
        return False
    # Location must match: prev destination = next origin
    if normalize_location(prev_trip.dest_code) != normalize_location(next_trip.origin_code):
        return False
    # Timing: enough turnaround (per bus type)
    min_turnaround = turnaround_map.get(prev_trip.bus_type, MIN_TURNAROUND_FALLBACK)
    gap = next_trip.departure - prev_trip.arrival
    if gap < min_turnaround:
        return False
    return True


def _group_trips(trips, turnaround_map):
    """Common setup: group trips by (date, bus_type), build compatibility edges."""
    if turnaround_map is None:
        turnaround_map = dict(MIN_TURNAROUND_DEFAULTS)
    groups = {}
    for t in trips:
        key = (t.date_str, t.bus_type)
        groups.setdefault(key, []).append(t)
    return groups, turnaround_map


def _build_rotations(group_trips, date_str, bus_type, chains, rotation_counter):
    """Convert list of trip-index chains into BusRotation objects."""
    rotations = []
    for chain in chains:
        rotation_counter += 1
        bus_id = f"{bus_type[:2].upper()}-{date_str.split()[0].upper()}-{rotation_counter:03d}"
        rotations.append(BusRotation(
            bus_id=bus_id,
            bus_type=bus_type,
            date_str=date_str,
            trips=[group_trips[i] for i in chain],
        ))
    return rotations, rotation_counter


# ---------------------------------------------------------------------------
# Algorithm 1: Greedy best-fit
# ---------------------------------------------------------------------------

def _optimize_greedy(group_trips, turnaround_map):
    """Greedy best-fit: assign each trip to the bus with shortest idle time."""
    group_trips.sort(key=lambda t: (t.departure, t.arrival))
    buses = []  # list of lists of trip indices

    for idx, trip in enumerate(group_trips):
        best_bus = None
        best_gap = float('inf')

        for bus in buses:
            last = group_trips[bus[-1]]
            if can_connect(last, trip, turnaround_map):
                gap = trip.departure - last.arrival
                if gap < best_gap:
                    best_gap = gap
                    best_bus = bus

        if best_bus is not None:
            best_bus.append(idx)
        else:
            buses.append([idx])

    return buses


# ---------------------------------------------------------------------------
# Algorithm 2: Maximum bipartite matching (Hopcroft-Karp)
#   Minimizes number of buses (= trips - max matching).
# ---------------------------------------------------------------------------

def _hopcroft_karp(adj, n_left, n_right):
    """
    Hopcroft-Karp algorithm for maximum bipartite matching.
    adj[u] = list of right-side nodes that left-side node u can match to.
    Returns (match_l, match_r) where match_l[u] = matched right node or -1.
    """
    from collections import deque

    match_l = [-1] * n_left
    match_r = [-1] * n_right

    def bfs():
        dist = [0] * n_left
        queue = deque()
        for u in range(n_left):
            if match_l[u] == -1:
                dist[u] = 0
                queue.append(u)
            else:
                dist[u] = float('inf')
        found = False
        while queue:
            u = queue.popleft()
            for v in adj[u]:
                w = match_r[v]
                if w == -1:
                    found = True
                elif dist[w] == float('inf'):
                    dist[w] = dist[u] + 1
                    queue.append(w)
        return found, dist

    def dfs(u, dist):
        for v in adj[u]:
            w = match_r[v]
            if w == -1 or (dist[w] == dist[u] + 1 and dfs(w, dist)):
                match_l[u] = v
                match_r[v] = u
                return True
        dist[u] = float('inf')
        return False

    while True:
        found, dist = bfs()
        if not found:
            break
        for u in range(n_left):
            if match_l[u] == -1:
                dfs(u, dist)

    return match_l, match_r


def _matching_to_chains(n, match_l):
    """Convert a matching into chains of trip indices."""
    # match_l[i] = j means trip i is followed by trip j
    matched_targets = set(v for v in match_l if v != -1)
    chains = []
    for i in range(n):
        if i not in matched_targets:
            # i is the start of a chain
            chain = [i]
            current = i
            while match_l[current] != -1:
                current = match_l[current]
                chain.append(current)
            chains.append(chain)
    return chains


def _optimize_matching(group_trips, turnaround_map):
    """Maximum bipartite matching via Hopcroft-Karp. Minimizes bus count."""
    group_trips.sort(key=lambda t: (t.departure, t.arrival))
    n = len(group_trips)

    # Build adjacency: trip i can be followed by trip j
    adj = [[] for _ in range(n)]
    for i in range(n):
        for j in range(i + 1, n):
            if can_connect(group_trips[i], group_trips[j], turnaround_map):
                adj[i].append(j)

    match_l, _ = _hopcroft_karp(adj, n, n)
    return _matching_to_chains(n, match_l)


# ---------------------------------------------------------------------------
# Algorithm 3: Minimum-cost maximum matching (successive shortest path)
#   Minimizes buses first, then minimizes total idle time.
# ---------------------------------------------------------------------------

def _optimize_mincost(group_trips, turnaround_map):
    """
    Min-cost max matching via successive shortest path (SPFA).
    Minimizes number of buses (primary) and total idle time (secondary).
    """
    from collections import deque

    group_trips.sort(key=lambda t: (t.departure, t.arrival))
    n = len(group_trips)

    # Build adjacency with costs (idle time)
    adj = [[] for _ in range(n)]
    cost_map = {}
    for i in range(n):
        for j in range(i + 1, n):
            if can_connect(group_trips[i], group_trips[j], turnaround_map):
                idle = group_trips[j].departure - group_trips[i].arrival
                adj[i].append(j)
                cost_map[(i, j)] = idle

    # Successive shortest path: find augmenting paths in order of increasing cost
    # Model as flow network with residual graph
    # match_l[i] = j means left i matched to right j
    match_l = [-1] * n
    match_r = [-1] * n

    def spfa_augment():
        """Find minimum-cost augmenting path using SPFA."""
        dist = [float('inf')] * n  # distance to right-side node j
        prev_l = [-1] * n  # previous left node on path to right j
        in_queue = [False] * n
        queue = deque()

        # Start from all unmatched left nodes
        # dist_left[u] = 0 for unmatched u
        dist_left = [float('inf')] * n
        for u in range(n):
            if match_l[u] == -1:
                dist_left[u] = 0
                # Relax edges from u
                for v in adj[u]:
                    c = cost_map[(u, v)]
                    if c < dist[v]:
                        dist[v] = c
                        prev_l[v] = u
                        if not in_queue[v]:
                            queue.append(v)
                            in_queue[v] = True

        # SPFA: relax through alternating paths
        while queue:
            v = queue.popleft()
            in_queue[v] = False

            # v is a right-side node, matched to w = match_r[v]
            w = match_r[v]
            if w == -1:
                continue  # free node, potential augmenting path end

            # Relax from w (go through the matched edge, then to new right nodes)
            new_dist_w = dist[v]  # cost to reach w through v's matched edge (0 cost)
            if new_dist_w < dist_left[w]:
                dist_left[w] = new_dist_w
                for v2 in adj[w]:
                    c = new_dist_w + cost_map[(w, v2)]
                    if c < dist[v2]:
                        dist[v2] = c
                        prev_l[v2] = w
                        if not in_queue[v2]:
                            queue.append(v2)
                            in_queue[v2] = True

        # Find the free right-side node with minimum distance
        best_v = -1
        best_d = float('inf')
        for v in range(n):
            if match_r[v] == -1 and dist[v] < best_d:
                best_d = dist[v]
                best_v = v

        if best_v == -1:
            return False  # no augmenting path

        # Trace back and augment
        v = best_v
        while v != -1:
            u = prev_l[v]
            old_v = match_l[u]
            match_l[u] = v
            match_r[v] = u
            v = old_v

        return True

    # Find all augmenting paths
    while spfa_augment():
        pass

    return _matching_to_chains(n, match_l)


# ---------------------------------------------------------------------------
# Main dispatcher
# ---------------------------------------------------------------------------

ALGORITHMS = {
    "greedy": ("Greedy best-fit", _optimize_greedy),
    "matching": ("Maximum bipartite matching (Hopcroft-Karp)", _optimize_matching),
    "mincost": ("Min-cost maximum matching (SPFA)", _optimize_mincost),
}


def optimize_rotations(trips: list, turnaround_map: dict = None,
                       algorithm: str = "greedy",
                       per_service: bool = False) -> list:
    """
    Optimize bus rotations using the specified algorithm.

    If per_service=True, only chains trips within the same service (Excel tab).
    If per_service=False, chains across all services (cross-tab optimization).
    """
    groups, turnaround_map = _group_trips(trips, turnaround_map)
    algo_name, algo_func = ALGORITHMS[algorithm]

    # If per_service, further split groups by service
    if per_service:
        new_groups = {}
        for (date_str, bus_type), group_trips in groups.items():
            by_svc = {}
            for t in group_trips:
                by_svc.setdefault(t.service, []).append(t)
            for svc, svc_trips in by_svc.items():
                new_groups[(date_str, bus_type, svc)] = svc_trips
        all_rotations = []
        rotation_counter = 0
        for key, group_trips in sorted(new_groups.items()):
            date_str, bus_type = key[0], key[1]
            chains = algo_func(group_trips, turnaround_map)
            rotations, rotation_counter = _build_rotations(
                group_trips, date_str, bus_type, chains, rotation_counter
            )
            all_rotations.extend(rotations)
        return all_rotations

    all_rotations = []
    rotation_counter = 0

    for (date_str, bus_type), group_trips in sorted(groups.items()):
        chains = algo_func(group_trips, turnaround_map)
        rotations, rotation_counter = _build_rotations(
            group_trips, date_str, bus_type, chains, rotation_counter
        )
        all_rotations.extend(rotations)

    return all_rotations


# ---------------------------------------------------------------------------
# Excel Output Generator
# ---------------------------------------------------------------------------

# Style constants
HEADER_FONT = Font(name="Calibri", bold=True, size=11)
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT_WHITE = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
SUBHEADER_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
BUS_HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
RESERVE_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
TIME_FORMAT = "HH:MM"


def apply_header_style(ws, row, col_start, col_end, fill=None, font=None):
    if fill is None:
        fill = HEADER_FILL
    if font is None:
        font = HEADER_FONT_WHITE
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = font
        cell.fill = fill
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")


def write_omloop_sheet(wb_out, rotations: list, reserves: list):
    """
    Tab 1: Busomloop - Transvision style per bus.
    Groups by date + bus_type, shows each bus's trip sequence.
    """
    # Group rotations by date
    by_date = {}
    for r in rotations:
        by_date.setdefault(r.date_str, []).append(r)

    for date_str, date_rotations in sorted(by_date.items()):
        # Group by bus type within date
        by_type = {}
        for r in date_rotations:
            by_type.setdefault(r.bus_type, []).append(r)

        for bus_type, type_rotations in sorted(by_type.items()):
            # Create sheet per date+type
            type_abbrev = {"Dubbeldekker": "DD", "Touringcar": "TC", "Lagevloerbus": "LVB", "Midi bus": "Midi", "Taxibus": "Taxi"}.get(bus_type, bus_type[:4])
            day_abbrev = date_str.split()[0] if date_str else "dag"
            sheet_name = f"Omloop {type_abbrev} {day_abbrev}"
            # Ensure unique sheet name (max 31 chars)
            sheet_name = sheet_name[:31]
            ws = wb_out.create_sheet(title=sheet_name)

            # Sort rotations by first departure
            type_rotations.sort(key=lambda r: r.start_time)

            # Layout: 3 buses per block of columns, like the example
            buses_per_row = 3
            cols_per_bus = 7  # Van, Naar, Van(t), Tot(t), Duur, Hold, spacer

            row = 1
            # Title
            ws.cell(row=row, column=1, value=f"Busomlopen {bus_type} - {date_str}")
            ws.cell(row=row, column=1).font = Font(name="Calibri", bold=True, size=14)
            row += 1

            # Summary line
            total_buses = len(type_rotations)
            ws.cell(row=row, column=1, value=f"Totaal bussen: {total_buses}")
            ws.cell(row=row, column=1).font = Font(name="Calibri", bold=True, size=11)
            row += 2

            # Process in blocks of buses_per_row
            for block_start in range(0, len(type_rotations), buses_per_row):
                block = type_rotations[block_start:block_start + buses_per_row]
                max_trips = max(len(b.trips) for b in block)

                # Bus headers
                for i, bus in enumerate(block):
                    base_col = 1 + i * cols_per_bus
                    ws.cell(row=row, column=base_col, value=bus.bus_id)
                    ws.cell(row=row, column=base_col).font = Font(bold=True, size=11, color="FFFFFF")
                    ws.cell(row=row, column=base_col).fill = BUS_HEADER_FILL
                    ws.cell(row=row, column=base_col + 1, value=bus.bus_type)
                    ws.cell(row=row, column=base_col + 1).font = Font(bold=True, size=11, color="FFFFFF")
                    ws.cell(row=row, column=base_col + 1).fill = BUS_HEADER_FILL
                    # Fill remaining header cols
                    for cc in range(base_col + 2, base_col + cols_per_bus - 1):
                        ws.cell(row=row, column=cc).fill = BUS_HEADER_FILL

                    # Dienst info
                    dienst_str = (f"{minutes_to_str(bus.start_time)} - {minutes_to_str(bus.end_time)} "
                                  f"({bus.total_dienst_minutes // 60}u{bus.total_dienst_minutes % 60:02d})")
                    ws.cell(row=row, column=base_col + 2, value=dienst_str)
                    ws.cell(row=row, column=base_col + 2).font = Font(bold=True, size=9, color="FFFFFF")
                    ws.cell(row=row, column=base_col + 2).fill = BUS_HEADER_FILL

                row += 1

                # Column headers
                headers = ["Van", "Naar", "Vertrek", "Aankomst", "Duur", "Wacht"]
                for i, bus in enumerate(block):
                    base_col = 1 + i * cols_per_bus
                    for j, h in enumerate(headers):
                        cell = ws.cell(row=row, column=base_col + j, value=h)
                        cell.font = HEADER_FONT
                        cell.fill = SUBHEADER_FILL
                        cell.border = THIN_BORDER
                        cell.alignment = Alignment(horizontal="center")
                row += 1

                # Trip rows
                for trip_idx in range(max_trips):
                    for i, bus in enumerate(block):
                        base_col = 1 + i * cols_per_bus
                        if trip_idx < len(bus.trips):
                            t = bus.trips[trip_idx]
                            ws.cell(row=row, column=base_col, value=t.origin_name)
                            ws.cell(row=row, column=base_col + 1, value=t.dest_name)
                            ws.cell(row=row, column=base_col + 2, value=minutes_to_time(t.departure))
                            ws.cell(row=row, column=base_col + 2).number_format = "HH:MM"
                            ws.cell(row=row, column=base_col + 3, value=minutes_to_time(t.arrival))
                            ws.cell(row=row, column=base_col + 3).number_format = "HH:MM"
                            dur = t.arrival - t.departure
                            ws.cell(row=row, column=base_col + 4, value=f"{dur // 60}:{dur % 60:02d}")

                            # Hold/wait time until next trip
                            if trip_idx < len(bus.trips) - 1:
                                next_t = bus.trips[trip_idx + 1]
                                hold = next_t.departure - t.arrival
                                ws.cell(row=row, column=base_col + 5, value=f"{hold // 60}:{hold % 60:02d}")

                            # Apply borders
                            for cc in range(base_col, base_col + 6):
                                ws.cell(row=row, column=cc).border = THIN_BORDER
                                ws.cell(row=row, column=cc).alignment = Alignment(horizontal="center")
                    row += 1

                # Subtotals for this block
                for i, bus in enumerate(block):
                    base_col = 1 + i * cols_per_bus
                    ws.cell(row=row, column=base_col, value="Ritten:")
                    ws.cell(row=row, column=base_col).font = Font(bold=True, size=9)
                    ws.cell(row=row, column=base_col + 1, value=len(bus.trips))
                    ws.cell(row=row, column=base_col + 2, value="Rijtijd:")
                    ws.cell(row=row, column=base_col + 2).font = Font(bold=True, size=9)
                    ride = bus.total_ride_minutes
                    ws.cell(row=row, column=base_col + 3, value=f"{ride // 60}:{ride % 60:02d}")
                    ws.cell(row=row, column=base_col + 4, value="Wacht:")
                    ws.cell(row=row, column=base_col + 4).font = Font(bold=True, size=9)
                    idle = bus.total_idle_minutes
                    ws.cell(row=row, column=base_col + 5, value=f"{idle // 60}:{idle % 60:02d}")
                row += 2

            # Reserve buses section for this date
            # Match reserves by checking day name substring in reserve's day field
            day_map = {"do": "donderdag", "vr": "vrijdag", "za": "zaterdag",
                       "zo": "zondag", "ma": "maandag"}
            full_day = day_map.get(day_abbrev, day_abbrev)
            date_reserves = [r for r in reserves
                            if full_day in r.day.lower() or day_abbrev in r.day.lower()[:2]]
            if date_reserves:
                ws.cell(row=row, column=1, value="Reservebussen")
                ws.cell(row=row, column=1).font = Font(bold=True, size=12)
                row += 1
                headers_r = ["Station", "Aantal", "Van", "Tot", "Opmerking"]
                for j, h in enumerate(headers_r):
                    cell = ws.cell(row=row, column=1 + j, value=h)
                    cell.font = HEADER_FONT
                    cell.fill = RESERVE_FILL
                    cell.border = THIN_BORDER
                row += 1
                for rb in date_reserves:
                    ws.cell(row=row, column=1, value=rb.station)
                    ws.cell(row=row, column=2, value=rb.count)
                    ws.cell(row=row, column=3, value=minutes_to_time(rb.start))
                    ws.cell(row=row, column=3).number_format = "HH:MM"
                    ws.cell(row=row, column=4, value=minutes_to_time(rb.end))
                    ws.cell(row=row, column=4).number_format = "HH:MM"
                    ws.cell(row=row, column=5, value=rb.remark)
                    for cc in range(1, 6):
                        ws.cell(row=row, column=cc).border = THIN_BORDER
                    row += 1

            # Column widths
            for c in range(1, 22):
                ws.column_dimensions[get_column_letter(c)].width = 16


def write_overzicht_sheet(wb_out, rotations: list, all_trips: list):
    """
    Tab 2: Overzicht - Overview of how trips connect per bus.
    Single sheet showing all trip chains.
    """
    ws = wb_out.create_sheet(title="Overzicht Ritsamenhang")

    row = 1
    ws.cell(row=row, column=1, value="Overzicht Ritsamenhang - Alle Busomlopen")
    ws.cell(row=row, column=1).font = Font(bold=True, size=14)
    row += 2

    # Headers
    headers = [
        "Bus ID", "Bustype", "Datum", "Rit #", "Busdienst", "Richting",
        "Van", "Naar", "Vertrek", "Aankomst", "Duur (min)",
        "Wachttijd (min)", "Busnr", "Snel/Stop"
    ]
    for j, h in enumerate(headers):
        cell = ws.cell(row=row, column=1 + j, value=h)
        cell.font = HEADER_FONT_WHITE
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")
    row += 1

    # Sort rotations by date, then bus type, then start time
    sorted_rotations = sorted(rotations, key=lambda r: (r.date_str, r.bus_type, r.start_time))

    alt_fill = PatternFill(start_color="E8F0FE", end_color="E8F0FE", fill_type="solid")

    for r_idx, rot in enumerate(sorted_rotations):
        use_fill = alt_fill if r_idx % 2 == 0 else None
        for t_idx, trip in enumerate(rot.trips):
            # Wachttijd
            wait = ""
            if t_idx < len(rot.trips) - 1:
                next_trip = rot.trips[t_idx + 1]
                wait = next_trip.departure - trip.arrival

            values = [
                rot.bus_id, rot.bus_type, rot.date_str,
                t_idx + 1, trip.service, trip.direction,
                trip.origin_name, trip.dest_name,
                minutes_to_time(trip.departure),
                minutes_to_time(trip.arrival),
                trip.arrival - trip.departure,
                wait if isinstance(wait, int) else "",
                trip.bus_nr, trip.snel_stop,
            ]
            for j, v in enumerate(values):
                cell = ws.cell(row=row, column=1 + j, value=v)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal="center")
                if use_fill:
                    cell.fill = use_fill
                if j in (8, 9):
                    cell.number_format = "HH:MM"
            row += 1

    # Column widths
    widths = [16, 14, 18, 8, 20, 10, 22, 22, 10, 10, 12, 14, 12, 12]
    for j, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(1 + j)].width = w

    # Freeze panes
    ws.freeze_panes = "A4"


def write_berekeningen_sheet(wb_out, rotations: list, all_trips: list, reserves: list, turnaround_map: dict = None, algorithm: str = "greedy"):
    """
    Tab 3: Berekeningen - Calculations and KPIs.
    """
    if turnaround_map is None:
        turnaround_map = dict(MIN_TURNAROUND_DEFAULTS)
    ws = wb_out.create_sheet(title="Berekeningen")

    row = 1
    ws.cell(row=row, column=1, value="Berekeningen Busomloop Optimalisatie")
    ws.cell(row=row, column=1).font = Font(bold=True, size=14)
    row += 2

    # --- Section 1: Summary per date + bus type ---
    ws.cell(row=row, column=1, value="1. Samenvatting per datum en bustype")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1

    sum_headers = [
        "Datum", "Bustype", "Aantal bussen", "Totaal ritten",
        "Totaal rijtijd (uur)", "Totaal wachttijd (uur)",
        "Totaal diensttijd (uur)", "Gem. ritten/bus",
        "Gem. diensttijd/bus (uur)", "Benutting (%)"
    ]
    for j, h in enumerate(sum_headers):
        cell = ws.cell(row=row, column=1 + j, value=h)
        cell.font = HEADER_FONT_WHITE
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    row += 1

    # Group rotations
    groups = {}
    for r in rotations:
        key = (r.date_str, r.bus_type)
        groups.setdefault(key, []).append(r)

    grand_buses = 0
    grand_trips = 0
    grand_ride = 0
    grand_idle = 0
    grand_dienst = 0

    for (date_str, bus_type), rots in sorted(groups.items()):
        n_buses = len(rots)
        n_trips = sum(len(r.trips) for r in rots)
        ride_min = sum(r.total_ride_minutes for r in rots)
        idle_min = sum(r.total_idle_minutes for r in rots)
        dienst_min = sum(r.total_dienst_minutes for r in rots)
        benutting = (ride_min / dienst_min * 100) if dienst_min > 0 else 0

        grand_buses += n_buses
        grand_trips += n_trips
        grand_ride += ride_min
        grand_idle += idle_min
        grand_dienst += dienst_min

        values = [
            date_str, bus_type, n_buses, n_trips,
            round(ride_min / 60, 1), round(idle_min / 60, 1),
            round(dienst_min / 60, 1),
            round(n_trips / n_buses, 1) if n_buses > 0 else 0,
            round(dienst_min / n_buses / 60, 1) if n_buses > 0 else 0,
            round(benutting, 1),
        ]
        for j, v in enumerate(values):
            cell = ws.cell(row=row, column=1 + j, value=v)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
        row += 1

    # Totals row
    benutting_total = (grand_ride / grand_dienst * 100) if grand_dienst > 0 else 0
    totals = [
        "TOTAAL", "", grand_buses, grand_trips,
        round(grand_ride / 60, 1), round(grand_idle / 60, 1),
        round(grand_dienst / 60, 1),
        round(grand_trips / grand_buses, 1) if grand_buses > 0 else 0,
        round(grand_dienst / grand_buses / 60, 1) if grand_buses > 0 else 0,
        round(benutting_total, 1),
    ]
    for j, v in enumerate(totals):
        cell = ws.cell(row=row, column=1 + j, value=v)
        cell.border = THIN_BORDER
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    row += 3

    # --- Section 2: Per-bus detail ---
    ws.cell(row=row, column=1, value="2. Detail per bus")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1

    detail_headers = [
        "Bus ID", "Bustype", "Datum", "Eerste rit",
        "Laatste rit", "Dienststart", "Diensteinde",
        "Diensttijd (uur)", "Rijtijd (uur)", "Wachttijd (uur)",
        "Aantal ritten", "Benutting (%)"
    ]
    for j, h in enumerate(detail_headers):
        cell = ws.cell(row=row, column=1 + j, value=h)
        cell.font = HEADER_FONT_WHITE
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    row += 1

    sorted_rots = sorted(rotations, key=lambda r: (r.date_str, r.bus_type, r.start_time))
    for rot in sorted_rots:
        if not rot.trips:
            continue
        first = rot.trips[0]
        last = rot.trips[-1]
        benutting = (rot.total_ride_minutes / rot.total_dienst_minutes * 100) if rot.total_dienst_minutes > 0 else 0
        values = [
            rot.bus_id, rot.bus_type, rot.date_str,
            f"{first.origin_name} -> {first.dest_name}",
            f"{last.origin_name} -> {last.dest_name}",
            minutes_to_time(rot.start_time),
            minutes_to_time(rot.end_time),
            round(rot.total_dienst_minutes / 60, 1),
            round(rot.total_ride_minutes / 60, 1),
            round(rot.total_idle_minutes / 60, 1),
            len(rot.trips),
            round(benutting, 1),
        ]
        for j, v in enumerate(values):
            cell = ws.cell(row=row, column=1 + j, value=v)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
            if j in (5, 6):
                cell.number_format = "HH:MM"
        row += 1

    row += 2

    # --- Section 3: Reservebussen ---
    ws.cell(row=row, column=1, value="3. Reservebussen")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1

    res_headers = ["Station", "Aantal", "Dag", "Van", "Tot", "Opmerking"]
    for j, h in enumerate(res_headers):
        cell = ws.cell(row=row, column=1 + j, value=h)
        cell.font = HEADER_FONT_WHITE
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")
    row += 1

    total_reserve = 0
    for rb in reserves:
        values = [rb.station, rb.count, rb.day,
                  minutes_to_time(rb.start), minutes_to_time(rb.end), rb.remark]
        for j, v in enumerate(values):
            cell = ws.cell(row=row, column=1 + j, value=v)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
            if j in (3, 4):
                cell.number_format = "HH:MM"
        total_reserve += rb.count
        row += 1

    ws.cell(row=row, column=1, value="Totaal reservebussen:")
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=2, value=total_reserve)
    ws.cell(row=row, column=2).font = Font(bold=True)
    row += 3

    # --- Section 4: Optimalisatie parameters ---
    ws.cell(row=row, column=1, value="4. Optimalisatie parameters & toelichting")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1

    params = [
        ("Algoritme", ALGORITHMS[algorithm][0]),
        ("Minimum keertijd", ", ".join(f"{bt}: {mins} min" for bt, mins in turnaround_map.items())),
        ("Doel", "Minimaliseer aantal bussen, daarna minimaliseer wachttijd"),
        ("Locatie-matching", "Bus eindlocatie moet gelijk zijn aan volgende rit startlocatie"),
        ("Bustype-constraint", "Bussen worden alleen ingezet op ritten met hetzelfde bustype"),
        ("Datum-constraint", "Bussen worden per dag apart geoptimaliseerd"),
        ("Multiplicity", "Ritten met 'Aantal bussen' > 1 worden uitgesplitst in aparte ritten"),
    ]
    for label, val in params:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=1).font = Font(bold=True)
        ws.cell(row=row, column=2, value=val)
        row += 1

    row += 2

    # --- Section 5: Niet-toegewezen ritten check ---
    ws.cell(row=row, column=1, value="5. Controle")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1

    total_input = len(all_trips)
    total_assigned = sum(len(r.trips) for r in rotations)
    ws.cell(row=row, column=1, value="Totaal ritten in invoer:")
    ws.cell(row=row, column=2, value=total_input)
    row += 1
    ws.cell(row=row, column=1, value="Totaal ritten toegewezen:")
    ws.cell(row=row, column=2, value=total_assigned)
    row += 1
    ws.cell(row=row, column=1, value="Niet-toegewezen ritten:")
    ws.cell(row=row, column=2, value=total_input - total_assigned)
    if total_input != total_assigned:
        ws.cell(row=row, column=2).font = Font(bold=True, color="FF0000")
    row += 1

    row += 2

    # --- Section 6: Algorithm examples ---
    row = _write_algo_examples(ws, row)

    # Column widths (don't override col A width set by _write_algo_examples)
    widths = [None, 16, 18, 28, 28, 14, 14, 16, 14, 16, 14, 14]
    for j, w in enumerate(widths):
        if w is not None and j < 12:
            ws.column_dimensions[get_column_letter(1 + j)].width = w


def write_businzet_sheet(wb_out, rotations: list, all_trips: list, reserves: list):
    """
    Tab 4: Overzicht Businzet - Matrix of services x dates with bus counts.
    Similar to 'Overzicht businzet.xlsx' but auto-generated with details.
    """
    ws = wb_out.create_sheet(title="Overzicht Businzet")

    row = 1
    ws.cell(row=row, column=1, value="Overzicht Businzet")
    ws.cell(row=row, column=1).font = Font(bold=True, size=14)
    row += 2

    # --- Section 1: Service x Date matrix ---
    ws.cell(row=row, column=1, value="1. Busdiensten per datum")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1

    # Collect unique dates and services from trips
    dates = sorted(set(t.date_str for t in all_trips))
    services = sorted(set(t.service for t in all_trips))

    # Build lookup: (service, date) -> {bus_type, trip_count, bus_count}
    from collections import defaultdict
    svc_date = defaultdict(lambda: {"trips": 0, "bus_types": set()})
    for t in all_trips:
        key = (t.service, t.date_str)
        svc_date[key]["trips"] += 1
        svc_date[key]["bus_types"].add(t.bus_type)

    # Count buses per service+date from rotations
    rot_svc_date = defaultdict(set)
    for rot in rotations:
        for trip in rot.trips:
            rot_svc_date[(trip.service, rot.date_str)].add(rot.bus_id)

    # Header row
    ws.cell(row=row, column=1, value="Busdienst")
    ws.cell(row=row, column=1).font = HEADER_FONT_WHITE
    ws.cell(row=row, column=1).fill = HEADER_FILL
    ws.cell(row=row, column=1).border = THIN_BORDER
    for j, date in enumerate(dates):
        col_base = 2 + j * 3
        for offset, header in enumerate(["Ritten", "Bussen", "Type"]):
            cell = ws.cell(row=row, column=col_base + offset, value=header)
            cell.font = HEADER_FONT_WHITE
            cell.fill = HEADER_FILL
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
    # Date labels above
    date_row = row - 1
    for j, date in enumerate(dates):
        col_base = 2 + j * 3
        cell = ws.cell(row=date_row, column=col_base, value=date)
        cell.font = Font(bold=True, size=11)
        cell.alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=date_row, start_column=col_base,
                       end_row=date_row, end_column=col_base + 2)
    row += 1

    type_abbrev = {"Dubbeldekker": "DD", "Touringcar": "TC", "Lagevloerbus": "LVB", "Midi bus": "Midi", "Taxibus": "Taxi"}
    alt_fill = PatternFill(start_color="E8F0FE", end_color="E8F0FE", fill_type="solid")

    for s_idx, service in enumerate(services):
        use_fill = alt_fill if s_idx % 2 == 0 else None
        cell = ws.cell(row=row, column=1, value=service)
        cell.font = Font(bold=True)
        cell.border = THIN_BORDER
        if use_fill:
            cell.fill = use_fill

        for j, date in enumerate(dates):
            col_base = 2 + j * 3
            key = (service, date)
            info = svc_date.get(key)
            if info and info["trips"] > 0:
                n_trips = info["trips"]
                n_buses = len(rot_svc_date.get(key, set()))
                types_str = ", ".join(type_abbrev.get(bt, bt) for bt in sorted(info["bus_types"]))

                ws.cell(row=row, column=col_base, value=n_trips)
                ws.cell(row=row, column=col_base + 1, value=n_buses)
                ws.cell(row=row, column=col_base + 2, value=types_str)
            else:
                ws.cell(row=row, column=col_base, value="-")
                ws.cell(row=row, column=col_base + 1, value="-")
                ws.cell(row=row, column=col_base + 2, value="-")

            for offset in range(3):
                c = ws.cell(row=row, column=col_base + offset)
                c.border = THIN_BORDER
                c.alignment = Alignment(horizontal="center")
                if use_fill:
                    c.fill = use_fill
        row += 1

    # Totals row
    ws.cell(row=row, column=1, value="TOTAAL")
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=1).border = THIN_BORDER
    for j, date in enumerate(dates):
        col_base = 2 + j * 3
        date_trips = sum(1 for t in all_trips if t.date_str == date)
        date_buses = len(set(r.bus_id for r in rotations if r.date_str == date))
        ws.cell(row=row, column=col_base, value=date_trips)
        ws.cell(row=row, column=col_base + 1, value=date_buses)
        for offset in range(3):
            c = ws.cell(row=row, column=col_base + offset)
            c.border = THIN_BORDER
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center")
    row += 3

    # --- Section 2: Buses per date + type summary ---
    ws.cell(row=row, column=1, value="2. Bussen per datum en bustype")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1

    bus_types_all = sorted(set(t.bus_type for t in all_trips))
    headers2 = ["Datum"] + [type_abbrev.get(bt, bt) for bt in bus_types_all] + ["Totaal", "Reserve"]
    for j, h in enumerate(headers2):
        cell = ws.cell(row=row, column=1 + j, value=h)
        cell.font = HEADER_FONT_WHITE
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")
    row += 1

    for date in dates:
        ws.cell(row=row, column=1, value=date)
        ws.cell(row=row, column=1).border = THIN_BORDER
        total_date = 0
        for bt_idx, bt in enumerate(bus_types_all):
            n = len([r for r in rotations if r.date_str == date and r.bus_type == bt])
            ws.cell(row=row, column=2 + bt_idx, value=n)
            ws.cell(row=row, column=2 + bt_idx).border = THIN_BORDER
            ws.cell(row=row, column=2 + bt_idx).alignment = Alignment(horizontal="center")
            total_date += n
        ws.cell(row=row, column=2 + len(bus_types_all), value=total_date)
        ws.cell(row=row, column=2 + len(bus_types_all)).border = THIN_BORDER
        ws.cell(row=row, column=2 + len(bus_types_all)).font = Font(bold=True)
        ws.cell(row=row, column=2 + len(bus_types_all)).alignment = Alignment(horizontal="center")

        # Reserve count for this date
        day_map = {"do": "donderdag", "vr": "vrijdag", "za": "zaterdag",
                   "zo": "zondag", "ma": "maandag"}
        day_abbrev = date.split()[0] if date else ""
        full_day = day_map.get(day_abbrev, day_abbrev)
        res_count = sum(r.count for r in reserves
                       if full_day in r.day.lower() or day_abbrev in r.day.lower()[:2])
        ws.cell(row=row, column=3 + len(bus_types_all), value=res_count)
        ws.cell(row=row, column=3 + len(bus_types_all)).border = THIN_BORDER
        ws.cell(row=row, column=3 + len(bus_types_all)).alignment = Alignment(horizontal="center")
        row += 1

    # Grand total
    ws.cell(row=row, column=1, value="TOTAAL")
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=1).border = THIN_BORDER
    grand = 0
    for bt_idx, bt in enumerate(bus_types_all):
        n = len([r for r in rotations if r.bus_type == bt])
        ws.cell(row=row, column=2 + bt_idx, value=n)
        ws.cell(row=row, column=2 + bt_idx).font = Font(bold=True)
        ws.cell(row=row, column=2 + bt_idx).border = THIN_BORDER
        ws.cell(row=row, column=2 + bt_idx).alignment = Alignment(horizontal="center")
        grand += n
    ws.cell(row=row, column=2 + len(bus_types_all), value=grand)
    ws.cell(row=row, column=2 + len(bus_types_all)).font = Font(bold=True)
    ws.cell(row=row, column=2 + len(bus_types_all)).border = THIN_BORDER
    ws.cell(row=row, column=2 + len(bus_types_all)).alignment = Alignment(horizontal="center")
    res_total = sum(r.count for r in reserves)
    ws.cell(row=row, column=3 + len(bus_types_all), value=res_total)
    ws.cell(row=row, column=3 + len(bus_types_all)).font = Font(bold=True)
    ws.cell(row=row, column=3 + len(bus_types_all)).border = THIN_BORDER
    ws.cell(row=row, column=3 + len(bus_types_all)).alignment = Alignment(horizontal="center")
    row += 3

    # --- Section 3: Diensttijden per busdienst ---
    ws.cell(row=row, column=1, value="3. Diensttijden per busdienst")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1

    headers3 = ["Busdienst", "Datum", "Bustype", "Eerste vertrek", "Laatste aankomst",
                "Diensttijd (uur)", "Ritten", "Bussen ingezet"]
    for j, h in enumerate(headers3):
        cell = ws.cell(row=row, column=1 + j, value=h)
        cell.font = HEADER_FONT_WHITE
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    row += 1

    for s_idx, service in enumerate(services):
        use_fill = alt_fill if s_idx % 2 == 0 else None
        for date in dates:
            key = (service, date)
            info = svc_date.get(key)
            if not info or info["trips"] == 0:
                continue
            svc_trips = [t for t in all_trips if t.service == service and t.date_str == date]
            if not svc_trips:
                continue
            first_dep = min(t.departure for t in svc_trips)
            last_arr = max(t.arrival for t in svc_trips)
            dienst_min = last_arr - first_dep
            n_buses = len(rot_svc_date.get(key, set()))
            types_str = ", ".join(type_abbrev.get(bt, bt) for bt in sorted(info["bus_types"]))

            values = [
                service, date, types_str,
                minutes_to_time(first_dep), minutes_to_time(last_arr),
                round(dienst_min / 60, 1), info["trips"], n_buses,
            ]
            for j, v in enumerate(values):
                cell = ws.cell(row=row, column=1 + j, value=v)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal="center")
                if use_fill:
                    cell.fill = use_fill
                if j in (3, 4):
                    cell.number_format = "HH:MM"
            row += 1

    # Column widths
    ws.column_dimensions[get_column_letter(1)].width = 20
    for c in range(2, 2 + len(dates) * 3 + 5):
        ws.column_dimensions[get_column_letter(c)].width = 14


def _write_algo_examples(ws, row):
    """Write simple, accessible algorithm examples to a worksheet."""
    ws.cell(row=row, column=1, value="6. Hoe werken de algoritmes? (voorbeelden)")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 2

    # Example scenario
    ws.cell(row=row, column=1, value="Stel: 4 ritten op dezelfde dag, allemaal Touringcar, minimale keertijd = 8 minuten")
    ws.cell(row=row, column=1).font = Font(italic=True)
    row += 1

    example_headers = ["Rit", "Van", "Naar", "Vertrek", "Aankomst"]
    for j, h in enumerate(example_headers):
        c = ws.cell(row=row, column=1 + j, value=h)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9E1F2")
        c.border = THIN_BORDER
    row += 1

    example_trips = [
        ("Rit 1", "Utrecht", "Ede", "06:00", "06:42"),
        ("Rit 2", "Ede", "Utrecht", "06:50", "07:32"),
        ("Rit 3", "Utrecht", "Ede", "07:00", "07:42"),
        ("Rit 4", "Ede", "Utrecht", "07:50", "08:32"),
    ]
    for vals in example_trips:
        for j, v in enumerate(vals):
            c = ws.cell(row=row, column=1 + j, value=v)
            c.border = THIN_BORDER
        row += 1
    row += 1

    # --- GREEDY ---
    ws.cell(row=row, column=1, value="A) Greedy best-fit (\"pak de eerste de beste\")")
    ws.cell(row=row, column=1).font = Font(bold=True, color="1F4E79")
    row += 1
    greedy_lines = [
        "Het algoritme loopt de ritten af op vertrektijd en koppelt elke rit aan de bus",
        "die het kortst stilstaat maar wel op de juiste plek staat.",
        "",
        "Stap 1: Rit 1 (Ut→Ed 06:00-06:42) → geen bus beschikbaar → Bus A",
        "Stap 2: Rit 2 (Ed→Ut 06:50-07:32) → Bus A staat in Ede, wacht 8 min → Bus A",
        "Stap 3: Rit 3 (Ut→Ed 07:00-07:42) → Bus A is onderweg → geen bus → Bus B",
        "Stap 4: Rit 4 (Ed→Ut 07:50-08:32) → Bus A in Ut (wacht 18 min), Bus B in Ede (wacht 8 min)",
        "        → Bus B (kortste wachttijd) → Bus B",
        "",
        "Resultaat: 2 bussen. Bus A: Rit 1→2 | Bus B: Rit 3→4",
        "Voordeel: Snel, werkt goed in de praktijk. Nadeel: vindt niet altijd het absolute minimum.",
    ]
    for line in greedy_lines:
        ws.cell(row=row, column=1, value=line)
        row += 1
    row += 1

    # --- MATCHING ---
    ws.cell(row=row, column=1, value="B) Bipartite matching (\"wiskundig optimaal minimum bussen\")")
    ws.cell(row=row, column=1).font = Font(bold=True, color="1F4E79")
    row += 1
    matching_lines = [
        "Het algoritme bouwt een netwerk van alle mogelijke koppelingen tussen ritten",
        "en vindt wiskundig het maximum aantal koppelingen → het minimum aantal bussen.",
        "",
        "Mogelijke koppelingen (aankomstlocatie = vertreklocatie, genoeg keertijd):",
        "  Rit 1 (aankomst Ede 06:42) → Rit 2 (vertrek Ede 06:50): gap 8 min ✓",
        "  Rit 1 (aankomst Ede 06:42) → Rit 4 (vertrek Ede 07:50): gap 68 min ✓",
        "  Rit 2 (aankomst Ut 07:32) → geen rit vertrekt uit Ut na 07:32",
        "  Rit 3 (aankomst Ede 07:42) → Rit 4 (vertrek Ede 07:50): gap 8 min ✓",
        "",
        "Maximale matching: Rit 1→Rit 2, Rit 3→Rit 4 (2 koppelingen)",
        "Bussen nodig = 4 ritten - 2 koppelingen = 2 bussen",
        "",
        "Resultaat: 2 bussen (gegarandeerd optimaal). Nadeel: optimaliseert niet op wachttijd.",
    ]
    for line in matching_lines:
        ws.cell(row=row, column=1, value=line)
        row += 1
    row += 1

    # --- MINCOST ---
    ws.cell(row=row, column=1, value="C) Min-cost matching (\"optimaal bussen + minimale wachttijd\")")
    ws.cell(row=row, column=1).font = Font(bold=True, color="1F4E79")
    row += 1
    mincost_lines = [
        "Zoals matching, maar bij gelijke aantallen bussen kiest het de verdeling",
        "met de minste totale wachttijd.",
        "",
        "Mogelijke oplossingen met 2 bussen:",
        "  Optie 1: Bus A: Rit 1→2 (wacht 8 min) | Bus B: Rit 3→4 (wacht 8 min)  → totaal 16 min wacht",
        "  Optie 2: Bus A: Rit 1→4 (wacht 68 min) | Bus B: Rit 3 | Bus C: Rit 2   → 3 bussen, slechter",
        "",
        "Min-cost kiest Optie 1: 2 bussen, 16 minuten totale wachttijd.",
        "Dit is het beste van beide werelden: minimum bussen EN minimum wachttijd.",
        "",
        "Resultaat: 2 bussen, 16 min wacht. Dit is de meest volledige optimalisatie.",
    ]
    for line in mincost_lines:
        ws.cell(row=row, column=1, value=line)
        row += 1

    ws.column_dimensions["A"].width = 90
    return row


def write_sensitivity_sheet(wb_out, all_trips: list, base_turnaround_map: dict,
                            algorithm: str = "greedy"):
    """
    Tab: Sensitiviteitsanalyse - shows impact of different turnaround times.
    For each bus type present in the data, varies turnaround time and shows bus count.
    """
    ws = wb_out.create_sheet(title="Sensitiviteit")
    row = 1
    ws.cell(row=row, column=1, value="Sensitiviteitsanalyse Keertijden")
    ws.cell(row=row, column=1).font = Font(bold=True, size=14)
    row += 2

    ws.cell(row=row, column=1, value="Wat als de minimale keertijd anders is? Hoeveel bussen zijn er dan nodig?")
    ws.cell(row=row, column=1).font = Font(italic=True)
    row += 1
    ws.cell(row=row, column=1, value=f"Algoritme: {ALGORITHMS[algorithm][0]}")
    row += 2

    # Get bus types present in data
    active_types = sorted(set(t.bus_type for t in all_trips))
    # Test values from 2 to max(base+4, 15)
    max_test = max(max(base_turnaround_map.get(bt, 8) for bt in active_types) + 4, 15)
    test_values = list(range(2, max_test + 1))

    # --- Section 1: Total buses per turnaround time variation ---
    ws.cell(row=row, column=1, value="1. Totaal bussen per keertijd-variatie (per bustype)")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1

    for bus_type in active_types:
        ws.cell(row=row, column=1, value=f"Bustype: {bus_type}")
        ws.cell(row=row, column=1).font = Font(bold=True)
        base_val = base_turnaround_map.get(bus_type, 8)
        ws.cell(row=row, column=2, value=f"(baseline: {base_val} min)")
        row += 1

        # Headers
        headers = ["Keertijd (min)", "Bussen nodig", "Verschil t.o.v. baseline", "Benutting (%)"]
        for j, h in enumerate(headers):
            c = ws.cell(row=row, column=1 + j, value=h)
            c.font = HEADER_FONT_WHITE
            c.fill = HEADER_FILL
            c.border = THIN_BORDER
            c.alignment = Alignment(horizontal="center")
        row += 1

        # Compute for each test value
        bt_trips = [t for t in all_trips if t.bus_type == bus_type]
        baseline_buses = None

        for tv in test_values:
            test_map = dict(base_turnaround_map)
            test_map[bus_type] = tv
            rots = optimize_rotations(all_trips, test_map, algorithm=algorithm)
            bt_rots = [r for r in rots if r.bus_type == bus_type]
            n_buses = len(bt_rots)
            ride = sum(r.total_ride_minutes for r in bt_rots)
            dienst = sum(r.total_dienst_minutes for r in bt_rots)
            benutting = (ride / dienst * 100) if dienst > 0 else 0

            if tv == base_val:
                baseline_buses = n_buses

            diff = n_buses - baseline_buses if baseline_buses is not None else 0
            diff_str = f"{diff:+d}" if baseline_buses is not None else ""

            is_base = (tv == base_val)
            vals = [f"{tv} min", n_buses, diff_str, round(benutting, 1)]
            for j, v in enumerate(vals):
                c = ws.cell(row=row, column=1 + j, value=v)
                c.border = THIN_BORDER
                c.alignment = Alignment(horizontal="center")
                if is_base:
                    c.font = Font(bold=True)
                    c.fill = PatternFill("solid", fgColor="E2EFDA")
            row += 1

        row += 2

    # --- Section 2: Detail of short turnarounds per route ---
    ws.cell(row=row, column=1, value="2. Overzicht korte keertijden per traject")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1

    ws.cell(row=row, column=1, value="Welke trajecten hebben korte keertijden? Zijn kortere keertijden realistisch?")
    ws.cell(row=row, column=1).font = Font(italic=True)
    row += 1

    route_headers = ["Dienst", "Bustype", "Min. keertijd (min)", "Locatie", "Voorbeeld aankomst", "Voorbeeld vertrek"]
    for j, h in enumerate(route_headers):
        c = ws.cell(row=row, column=1 + j, value=h)
        c.font = HEADER_FONT_WHITE
        c.fill = HEADER_FILL
        c.border = THIN_BORDER
        c.alignment = Alignment(horizontal="center")
    row += 1

    # Find min gap per service with example
    by_service = {}
    for t in all_trips:
        by_service.setdefault(t.service, []).append(t)

    service_gaps = []
    for service, svc_trips in sorted(by_service.items()):
        bus_type = svc_trips[0].bus_type
        arrivals_by_loc = {}
        departures_by_loc = {}
        for t in svc_trips:
            dest_loc = normalize_location(t.dest_code)
            orig_loc = normalize_location(t.origin_code)
            arrivals_by_loc.setdefault(dest_loc, []).append(t)
            departures_by_loc.setdefault(orig_loc, []).append(t)

        best_gap = None
        best_example = None
        for loc, arrs in arrivals_by_loc.items():
            deps = departures_by_loc.get(loc, [])
            deps_sorted = sorted(deps, key=lambda x: x.departure)
            for a in arrs:
                for d in deps_sorted:
                    g = d.departure - a.arrival
                    if g >= MIN_TURNAROUND_FLOOR:
                        if best_gap is None or g < best_gap:
                            best_gap = g
                            best_example = (loc, a, d)
                        break

        if best_gap is not None and best_example is not None:
            loc, a_trip, d_trip = best_example
            service_gaps.append((best_gap, service, bus_type, loc, a_trip, d_trip))

    service_gaps.sort(key=lambda x: x[0])
    for gap, service, bus_type, loc, a_trip, d_trip in service_gaps:
        vals = [
            service, bus_type, gap, loc,
            f"{minutes_to_str(a_trip.arrival)} ({a_trip.origin_name}→{a_trip.dest_name})",
            f"{minutes_to_str(d_trip.departure)} ({d_trip.origin_name}→{d_trip.dest_name})",
        ]
        for j, v in enumerate(vals):
            c = ws.cell(row=row, column=1 + j, value=v)
            c.border = THIN_BORDER
            c.alignment = Alignment(horizontal="center" if j in (2,) else "left")
        row += 1

    # Column widths
    for j, w in enumerate([28, 16, 18, 18, 40, 40]):
        ws.column_dimensions[get_column_letter(1 + j)].width = w


def generate_output(rotations: list, all_trips: list, reserves: list, output_file: str,
                    turnaround_map: dict = None, algorithm: str = "greedy",
                    include_sensitivity: bool = False):
    """Generate the complete output Excel workbook."""
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Tab 1: Busomloop (Transvision-stijl)
    write_omloop_sheet(wb, rotations, reserves)

    # Tab 2: Overzicht ritsamenhang
    write_overzicht_sheet(wb, rotations, all_trips)

    # Tab 3: Berekeningen
    write_berekeningen_sheet(wb, rotations, all_trips, reserves, turnaround_map, algorithm)

    # Tab 4: Overzicht Businzet
    write_businzet_sheet(wb, rotations, all_trips, reserves)

    # Tab 5 (optional): Sensitiviteitsanalyse
    if include_sensitivity:
        write_sensitivity_sheet(wb, all_trips, turnaround_map, algorithm)

    wb.save(output_file)
    return output_file


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Busomloop Optimizer - Genereert optimale busomlopen uit NS TVV dienstregeling"
    )
    parser.add_argument(
        "input_file",
        help="Invoer Excel bestand (Bijlage J casus busdiensten)",
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Uitvoer Excel bestand (standaard: busomloop_output.xlsx)",
    )
    parser.add_argument(
        "--algoritme", "-a",
        choices=list(ALGORITHMS.keys()),
        default="greedy",
        help="Optimalisatie-algoritme: greedy (snel, heuristisch), "
             "matching (optimaal min. bussen), mincost (optimaal min. bussen + min. wachttijd). "
             "Standaard: greedy",
    )
    parser.add_argument(
        "--keer-dd",
        type=int,
        default=None,
        help=f"Keertijd dubbeldekker in minuten (standaard: auto-detect uit data)",
    )
    parser.add_argument(
        "--keer-tc",
        type=int,
        default=None,
        help=f"Keertijd touringcar in minuten (standaard: auto-detect uit data)",
    )
    parser.add_argument(
        "--keer-lvb",
        type=int,
        default=None,
        help=f"Keertijd lagevloerbus/gelede bus in minuten (standaard: auto-detect uit data)",
    )
    parser.add_argument(
        "--keer-midi",
        type=int,
        default=None,
        help=f"Keertijd midi bus in minuten (standaard: auto-detect uit data)",
    )
    parser.add_argument(
        "--keer-taxi",
        type=int,
        default=None,
        help=f"Keertijd taxibus in minuten (standaard: auto-detect uit data)",
    )
    args = parser.parse_args()

    if args.output is None:
        args.output = "busomloop_output"

    # Strip .xlsx if user provided it (we'll add suffixes)
    output_base = args.output.replace(".xlsx", "")

    print(f"Busomloop Optimizer")
    print(f"{'='*60}")
    print(f"Invoer:        {args.input_file}")
    print(f"Uitvoer:       {output_base}_baseline.xlsx + {output_base}_gecombineerd.xlsx")
    print()

    # ===== PARSE =====
    print("Stap 1: Invoer parsen...")
    all_trips, reserves, sheet_names = parse_all_sheets(args.input_file)
    print(f"  {len(sheet_names) - 1} dienstbladen gevonden")
    print(f"  {len(all_trips)} ritten geparsed (inclusief multipliciteit)")
    print(f"  {len(reserves)} reservebus-regels gevonden")

    by_type = {}
    for t in all_trips:
        by_type.setdefault(t.bus_type, []).append(t)
    for bt, trips in sorted(by_type.items()):
        print(f"    {bt}: {len(trips)} ritten")
    print()

    # ===== DETECT TURNAROUND TIMES (within-service = baseline) =====
    print("Stap 2: Keertijden bepalen (per tabblad, zonder combinaties)...")
    baseline_turnaround = detect_turnaround_times(all_trips, within_service_only=True)

    # Apply CLI overrides
    cli_overrides = {
        "Dubbeldekker": args.keer_dd,
        "Touringcar": args.keer_tc,
        "Lagevloerbus": args.keer_lvb,
        "Midi bus": args.keer_midi,
        "Taxibus": args.keer_taxi,
    }
    for bt, val in cli_overrides.items():
        if val is not None:
            baseline_turnaround[bt] = val

    for bt in MIN_TURNAROUND_DEFAULTS:
        if bt not in baseline_turnaround:
            baseline_turnaround[bt] = MIN_TURNAROUND_DEFAULTS[bt]

    print(f"  Baseline keertijden (bepaald per tabblad):")
    for bt, mins in sorted(baseline_turnaround.items()):
        source = "handmatig" if cli_overrides.get(bt) is not None else "uit data"
        print(f"    {bt:20s} {mins:3d} min  ({source})")

    # Show per-service detail
    svc_turnarounds = detect_turnaround_per_service(all_trips)
    print(f"\n  Detail per dienst:")
    for svc, (bt, gap) in sorted(svc_turnarounds.items(), key=lambda x: x[1][1]):
        print(f"    {svc:30s} ({bt:15s})  keertijd {gap:3d} min")
    print()

    # ===== BASELINE OUTPUT (per-service, no cross-sheet combining) =====
    algo_name = ALGORITHMS[args.algoritme][0]
    print(f"Stap 3a: Baseline busomlopen (per tabblad, {algo_name})...")
    baseline_rotations = optimize_rotations(all_trips, baseline_turnaround, algorithm=args.algoritme, per_service=True)
    n_baseline = len(baseline_rotations)
    print(f"  {n_baseline} busomlopen (baseline)")

    baseline_file = f"{output_base}_baseline.xlsx"
    generate_output(baseline_rotations, all_trips, reserves, baseline_file,
                    baseline_turnaround, args.algoritme)
    print(f"  Opgeslagen: {baseline_file}")
    print()

    # ===== COMBINED OUTPUT (cross-sheet, same turnaround times + sensitivity) =====
    print(f"Stap 3b: Gecombineerde busomlopen (cross-tabblad, {algo_name})...")
    combined_rotations = optimize_rotations(all_trips, baseline_turnaround, algorithm=args.algoritme)
    n_combined = len(combined_rotations)
    print(f"  {n_combined} busomlopen (gecombineerd)")

    # Summary per date+type
    groups = {}
    for r in combined_rotations:
        key = (r.date_str, r.bus_type)
        groups.setdefault(key, []).append(r)
    for (date_str, bus_type), rots in sorted(groups.items()):
        ride = sum(r.total_ride_minutes for r in rots)
        dienst = sum(r.total_dienst_minutes for r in rots)
        benutting = (ride / dienst * 100) if dienst > 0 else 0
        print(f"    {date_str} / {bus_type}: {len(rots)} bussen, "
              f"{sum(len(r.trips) for r in rots)} ritten, "
              f"benutting {benutting:.0f}%")

    combined_file = f"{output_base}_gecombineerd.xlsx"
    print(f"\n  Sensitiviteitsanalyse genereren...")
    generate_output(combined_rotations, all_trips, reserves, combined_file,
                    baseline_turnaround, args.algoritme, include_sensitivity=True)
    print(f"  Opgeslagen: {combined_file}")
    print()

    # ===== FINAL SUMMARY =====
    def _summary(label, rots):
        n = len(rots)
        trips = sum(len(r.trips) for r in rots)
        ride = sum(r.total_ride_minutes for r in rots)
        dienst = sum(r.total_dienst_minutes for r in rots)
        benut = (ride / dienst * 100) if dienst > 0 else 0
        print(f"  {label}:")
        print(f"    Bussen:       {n}")
        print(f"    Ritten:       {trips}")
        print(f"    Rijtijd:      {ride // 60}u{ride % 60:02d}")
        print(f"    Diensttijd:   {dienst // 60}u{dienst % 60:02d}")
        print(f"    Benutting:    {benut:.1f}%")

    print(f"Vergelijking:")
    _summary("Baseline (per tabblad)", baseline_rotations)
    _summary("Gecombineerd (cross-tabblad)", combined_rotations)
    print(f"  Reservebussen:    {sum(r.count for r in reserves)}")
    print()
    print("Klaar!")


if __name__ == "__main__":
    main()
