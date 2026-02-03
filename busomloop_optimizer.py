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


def detect_turnaround_times(trips: list) -> dict:
    """
    Auto-detect minimum turnaround time per bus type from the trip data.

    For each bus type, groups trips by (date, location) and finds the smallest
    gap between an arriving trip and a departing trip at the same location.
    That smallest gap is the intended turnaround time for that bus type.

    Returns dict {bus_type: minutes}, with a floor of MIN_TURNAROUND_FLOOR.
    """
    # Group arrivals and departures by (bus_type, date, normalized_location)
    arrivals = {}   # key -> sorted list of arrival times
    departures = {} # key -> sorted list of departure times

    for t in trips:
        dest_loc = normalize_location(t.dest_code)
        orig_loc = normalize_location(t.origin_code)
        arr_key = (t.bus_type, t.date_str, dest_loc)
        dep_key = (t.bus_type, t.date_str, orig_loc)
        arrivals.setdefault(arr_key, []).append(t.arrival)
        departures.setdefault(dep_key, []).append(t.departure)

    # For each bus type, find minimum gap between any arrival and subsequent departure
    min_gap_per_type = {}

    for arr_key, arr_times in arrivals.items():
        bus_type, date_str, location = arr_key
        dep_key = (bus_type, date_str, location)
        if dep_key not in departures:
            continue

        dep_times = sorted(departures[dep_key])
        for arr_t in arr_times:
            # Find departures that happen after this arrival (binary search style)
            for dep_t in dep_times:
                gap = dep_t - arr_t
                if gap >= MIN_TURNAROUND_FLOOR:
                    if bus_type not in min_gap_per_type or gap < min_gap_per_type[bus_type]:
                        min_gap_per_type[bus_type] = gap
                    break  # smallest valid gap for this arrival found

    # Build result with fallback for types not found in data
    result = {}
    for bus_type in set(t.bus_type for t in trips):
        if bus_type in min_gap_per_type:
            result[bus_type] = min_gap_per_type[bus_type]
        else:
            result[bus_type] = MIN_TURNAROUND_FALLBACK

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


def optimize_rotations(trips: list, turnaround_map: dict = None) -> list:
    """
    Greedy best-fit algorithm to chain trips into bus rotations.
    Groups by (date, bus_type) and minimizes number of buses.
    """
    if turnaround_map is None:
        turnaround_map = dict(MIN_TURNAROUND_DEFAULTS)

    # Group trips by date + bus type
    groups = {}
    for t in trips:
        key = (t.date_str, t.bus_type)
        groups.setdefault(key, []).append(t)

    all_rotations = []
    rotation_counter = 0

    for (date_str, bus_type), group_trips in sorted(groups.items()):
        # Sort trips by departure time
        group_trips.sort(key=lambda t: (t.departure, t.arrival))

        # Active buses: list of BusRotation, each with last trip info
        active_buses = []

        for trip in group_trips:
            # Find best available bus: the one whose last trip ends latest
            # but still allows connection (minimize idle)
            best_bus = None
            best_gap = float('inf')

            for bus in active_buses:
                last = bus.trips[-1]
                if can_connect(last, trip, turnaround_map):
                    gap = trip.departure - last.arrival
                    if gap < best_gap:
                        best_gap = gap
                        best_bus = bus

            if best_bus is not None:
                best_bus.trips.append(trip)
            else:
                # Need a new bus
                rotation_counter += 1
                bus_id = f"{bus_type[:2].upper()}-{date_str.split()[0].upper()}-{rotation_counter:03d}"
                new_bus = BusRotation(
                    bus_id=bus_id,
                    bus_type=bus_type,
                    date_str=date_str,
                    trips=[trip],
                )
                active_buses.append(new_bus)

        all_rotations.extend(active_buses)

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


def write_berekeningen_sheet(wb_out, rotations: list, all_trips: list, reserves: list, turnaround_map: dict = None):
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
        ("Algoritme", "Greedy best-fit bus chaining"),
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

    # Column widths
    widths = [22, 16, 18, 28, 28, 14, 14, 16, 14, 16, 14, 14]
    for j, w in enumerate(widths):
        if j < 12:
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


def generate_output(rotations: list, all_trips: list, reserves: list, output_file: str, turnaround_map: dict = None):
    """Generate the complete output Excel workbook."""
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Tab 1: Busomloop (Transvision-stijl)
    write_omloop_sheet(wb, rotations, reserves)

    # Tab 2: Overzicht ritsamenhang
    write_overzicht_sheet(wb, rotations, all_trips)

    # Tab 3: Berekeningen
    write_berekeningen_sheet(wb, rotations, all_trips, reserves, turnaround_map)

    # Tab 4: Overzicht Businzet
    write_businzet_sheet(wb, rotations, all_trips, reserves)

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
        args.output = "busomloop_output.xlsx"

    print(f"Busomloop Optimizer")
    print(f"{'='*50}")
    print(f"Invoer:        {args.input_file}")
    print(f"Uitvoer:       {args.output}")
    print()

    # Parse
    print("Stap 1: Invoer parsen...")
    all_trips, reserves, sheet_names = parse_all_sheets(args.input_file)
    print(f"  {len(sheet_names) - 1} dienstbladen gevonden")
    print(f"  {len(all_trips)} ritten geparsed (inclusief multipliciteit)")
    print(f"  {len(reserves)} reservebus-regels gevonden")

    # Group trips by type for summary
    by_type = {}
    for t in all_trips:
        by_type.setdefault(t.bus_type, []).append(t)
    for bt, trips in sorted(by_type.items()):
        print(f"    {bt}: {len(trips)} ritten")
    print()

    # Determine turnaround times: auto-detect from data, then apply overrides
    print("Stap 1b: Keertijden bepalen...")
    auto_detected = detect_turnaround_times(all_trips)
    turnaround_map = dict(auto_detected)

    # Apply any explicit CLI overrides
    cli_overrides = {
        "Dubbeldekker": args.keer_dd,
        "Touringcar": args.keer_tc,
        "Lagevloerbus": args.keer_lvb,
        "Midi bus": args.keer_midi,
        "Taxibus": args.keer_taxi,
    }
    for bt, val in cli_overrides.items():
        if val is not None:
            turnaround_map[bt] = val

    # Ensure all known types have a value (fallback for types not in data)
    for bt in MIN_TURNAROUND_DEFAULTS:
        if bt not in turnaround_map:
            turnaround_map[bt] = MIN_TURNAROUND_DEFAULTS[bt]

    print(f"  Keertijden (min {MIN_TURNAROUND_FLOOR} min):")
    for bt, mins in sorted(turnaround_map.items()):
        source = "auto-detect" if cli_overrides.get(bt) is None and bt in auto_detected else "handmatig"
        print(f"    {bt:20s} {mins:3d} min  ({source})")
    print()

    # Optimize
    print("Stap 2: Busomlopen optimaliseren...")
    rotations = optimize_rotations(all_trips, turnaround_map)
    print(f"  {len(rotations)} busomlopen gegenereerd")

    # Summary per date+type
    groups = {}
    for r in rotations:
        key = (r.date_str, r.bus_type)
        groups.setdefault(key, []).append(r)
    for (date_str, bus_type), rots in sorted(groups.items()):
        ride = sum(r.total_ride_minutes for r in rots)
        dienst = sum(r.total_dienst_minutes for r in rots)
        benutting = (ride / dienst * 100) if dienst > 0 else 0
        print(f"    {date_str} / {bus_type}: {len(rots)} bussen, "
              f"{sum(len(r.trips) for r in rots)} ritten, "
              f"benutting {benutting:.0f}%")
    print()

    # Generate output
    print("Stap 3: Uitvoer genereren...")
    output = generate_output(rotations, all_trips, reserves, args.output, turnaround_map)
    print(f"  Uitvoer opgeslagen: {output}")
    print()

    # Final summary
    total_buses = len(rotations)
    total_trips = sum(len(r.trips) for r in rotations)
    total_ride = sum(r.total_ride_minutes for r in rotations)
    total_dienst = sum(r.total_dienst_minutes for r in rotations)
    print(f"Resultaat:")
    print(f"  Totaal bussen:      {total_buses}")
    print(f"  Totaal ritten:      {total_trips}")
    print(f"  Totaal rijtijd:     {total_ride // 60}u{total_ride % 60:02d}")
    print(f"  Totaal diensttijd:  {total_dienst // 60}u{total_dienst % 60:02d}")
    print(f"  Gem. benutting:     {total_ride / total_dienst * 100:.1f}%" if total_dienst > 0 else "  Gem. benutting:     N/A")
    print(f"  Reservebussen:      {sum(r.count for r in reserves)}")
    print()
    print("Klaar!")


if __name__ == "__main__":
    main()
