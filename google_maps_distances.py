"""
Google Maps Distance Matrix - Deadhead & Trip Validation
=========================================================
Fetches driving distances/times between all bus station locations using the
Google Maps Distance Matrix API. Produces:

1. A deadhead matrix (travel time in minutes between all station pairs)
2. A validation report comparing scheduled trip durations to Google Maps times
3. A traffic sensitivity analysis (where buffers are thin)

Usage:
    python google_maps_distances.py [--key API_KEY] [--output distances.xlsx]
    python google_maps_distances.py --from-cache deadhead_matrix.json --validate output.xlsx
"""

import argparse
import datetime
import json
import sys
import time
from pathlib import Path

import requests

# ---------------------------------------------------------------------------
# Canonical locations → Google Maps searchable addresses (train stations)
# ---------------------------------------------------------------------------
STATION_ADDRESSES = {
    "amersfoort":        "Station Amersfoort Centraal, Amersfoort, Nederland",
    "arnhem":            "Station Arnhem Centraal, Arnhem, Nederland",
    "bilthoven":         "Station Bilthoven, Bilthoven, Nederland",
    "breukelen":         "Station Breukelen, Breukelen, Nederland",
    "den_dolder":        "Station Den Dolder, Den Dolder, Nederland",
    "driebergen":        "Station Driebergen-Zeist, Driebergen, Nederland",
    "ede":               "Station Ede-Wageningen, Ede, Nederland",
    "geldermalsen":      "Station Geldermalsen, Geldermalsen, Nederland",
    "hilversum":         "Station Hilversum, Hilversum, Nederland",
    "houten":            "Station Houten, Houten, Nederland",
    "houten_vinex":      "Station Houten Castellum, Houten, Nederland",
    "maarn":             "Station Maarn, Maarn, Nederland",
    "rhenen":            "Station Rhenen, Rhenen, Nederland",
    "utrecht":           "Station Utrecht Centraal, Utrecht, Nederland",
    "utrecht_overvecht": "Station Utrecht Overvecht, Utrecht, Nederland",
    "veenendaal":        "Station Veenendaal-De Klomp, Veenendaal, Nederland",
    "veenendaal_centrum":"Station Veenendaal Centrum, Veenendaal, Nederland",
    "veenendaal_klomp":  "Station Veenendaal-De Klomp, Veenendaal, Nederland",
    "veenendaal_west":   "Station Veenendaal West, Veenendaal, Nederland",
    "woerden":           "Station Woerden, Woerden, Nederland",
}

# Map display names (from optimizer output) back to canonical locations
DISPLAY_TO_CANONICAL = {
    "arnhem centraal": "arnhem",
    "ede-wageningen": "ede",
    "utrecht centraal": "utrecht",
    "utrecht overvecht": "utrecht_overvecht",
    "driebergen-zeist": "driebergen",
    "veenendaal-de klomp": "veenendaal_klomp",
    "veenendaal centrum": "veenendaal_centrum",
    "veenendaal west": "veenendaal_west",
    "amersfoort centraal": "amersfoort",
    "houten castellum": "houten_vinex",
    "houten": "houten",
    "bilthoven": "bilthoven",
    "breukelen": "breukelen",
    "den dolder": "den_dolder",
    "geldermalsen": "geldermalsen",
    "hilversum": "hilversum",
    "maarn": "maarn",
    "rhenen": "rhenen",
    "woerden": "woerden",
}

API_URL = "https://maps.googleapis.com/maps/api/distancematrix/json"


def normalize_display_name(name: str) -> str:
    """Convert a display station name to canonical location."""
    key = name.strip().lower()
    return DISPLAY_TO_CANONICAL.get(key, key)


def fetch_distance_matrix(api_key: str, locations: list[str],
                          batch_size: int = 10) -> dict:
    """Fetch full NxN distance matrix from Google Maps.

    Returns dict: {(origin, dest): {"distance_m": int, "duration_s": int,
                                     "duration_min": float, "distance_km": float}}
    """
    addresses = [STATION_ADDRESSES[loc] for loc in locations]
    n = len(locations)
    matrix = {}

    for i_start in range(0, n, batch_size):
        i_end = min(i_start + batch_size, n)
        origins = "|".join(addresses[i_start:i_end])
        origin_locs = locations[i_start:i_end]

        for j_start in range(0, n, batch_size):
            j_end = min(j_start + batch_size, n)
            destinations = "|".join(addresses[j_start:j_end])
            dest_locs = locations[j_start:j_end]

            params = {
                "origins": origins,
                "destinations": destinations,
                "key": api_key,
                "mode": "driving",
                "language": "nl",
            }

            print(f"  Fetching {len(origin_locs)}x{len(dest_locs)} "
                  f"({origin_locs[0]}..{origin_locs[-1]} -> "
                  f"{dest_locs[0]}..{dest_locs[-1]}) ...")

            resp = requests.get(API_URL, params=params, timeout=30)
            resp.raise_for_status()
            data = resp.json()

            if data["status"] != "OK":
                print(f"  ERROR: API returned status {data['status']}")
                print(f"  {data.get('error_message', '')}")
                sys.exit(1)

            for ri, row in enumerate(data["rows"]):
                for ci, element in enumerate(row["elements"]):
                    o = origin_locs[ri]
                    d = dest_locs[ci]
                    if element["status"] == "OK":
                        matrix[(o, d)] = {
                            "distance_m": element["distance"]["value"],
                            "duration_s": element["duration"]["value"],
                            "duration_min": round(element["duration"]["value"] / 60, 1),
                            "distance_km": round(element["distance"]["value"] / 1000, 1),
                        }
                    else:
                        print(f"  WARNING: No route {o} -> {d}: {element['status']}")
                        matrix[(o, d)] = {
                            "distance_m": None, "duration_s": None,
                            "duration_min": None, "distance_km": None,
                        }

            time.sleep(0.2)

    return matrix


def save_deadhead_json(matrix: dict, locations: list[str], output_file: str):
    """Save deadhead matrix as JSON for the optimizer to import."""
    deadhead = {}
    for o in locations:
        deadhead[o] = {}
        for d in locations:
            entry = matrix.get((o, d))
            if entry and entry["duration_min"] is not None:
                deadhead[o][d] = entry["duration_min"]
    with open(output_file, "w") as f:
        json.dump(deadhead, f, indent=2, ensure_ascii=False)
    print(f"Deadhead JSON saved to {output_file}")


def load_matrix_from_cache(cache_file: str) -> dict:
    """Load distance matrix from cached JSON."""
    with open(cache_file) as f:
        cached = json.load(f)
    matrix = {}
    for o, dests in cached.items():
        for d, val in dests.items():
            # Cache stores just the duration_min value
            if isinstance(val, (int, float)):
                matrix[(o, d)] = {"duration_min": val, "distance_km": None,
                                  "duration_s": None, "distance_m": None}
            else:
                matrix[(o, d)] = val
    return matrix


# ---------------------------------------------------------------------------
# Trip validation from optimizer output
# ---------------------------------------------------------------------------

def load_trips_from_output(output_file: str) -> list[dict]:
    """Read trips from optimizer output 'Overzicht Ritsamenhang' sheet.

    Returns list of dicts: {bus_id, origin, dest, origin_loc, dest_loc,
                            departure, arrival, duration_min, service, direction}
    """
    import openpyxl

    wb = openpyxl.load_workbook(output_file, data_only=True)

    # Find the ritsamenhang sheet
    ws = None
    for name in wb.sheetnames:
        if "ritsamenhang" in name.lower():
            ws = wb[name]
            break

    if not ws:
        print(f"Could not find 'Ritsamenhang' sheet in {output_file}")
        wb.close()
        return []

    # Find header row (contains "Bus ID", "Van", "Naar", etc.)
    header_map = {}
    header_row = None
    for row_idx in range(1, min(10, ws.max_row + 1)):
        val = ws.cell(row_idx, 1).value
        if val and str(val).strip().lower() == "bus id":
            header_row = row_idx
            for c in range(1, ws.max_column + 1):
                h = ws.cell(row_idx, c).value
                if h:
                    header_map[str(h).strip().lower()] = c
            break

    if not header_row:
        print("Could not find header row in Ritsamenhang sheet")
        wb.close()
        return []

    # Map columns
    col_bus = header_map.get("bus id", 1)
    col_type = header_map.get("bustype", 2)
    col_service = header_map.get("busdienst", 5)
    col_direction = header_map.get("richting", 6)
    col_origin = header_map.get("van", 7)
    col_dest = header_map.get("naar", 8)
    col_dep = header_map.get("vertrek", 9)
    col_arr = header_map.get("aankomst", 10)
    col_dur = header_map.get("duur (min)", 11)

    trips = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        bus_id = ws.cell(row_idx, col_bus).value
        if not bus_id:
            continue

        origin = ws.cell(row_idx, col_origin).value or ""
        dest = ws.cell(row_idx, col_dest).value or ""
        dep = ws.cell(row_idx, col_dep).value
        arr = ws.cell(row_idx, col_arr).value
        dur = ws.cell(row_idx, col_dur).value
        service = ws.cell(row_idx, col_service).value or ""
        direction = ws.cell(row_idx, col_direction).value or ""
        bus_type = ws.cell(row_idx, col_type).value or ""

        # Convert times to minutes
        dep_min = None
        arr_min = None
        if isinstance(dep, datetime.time):
            dep_min = dep.hour * 60 + dep.minute
        if isinstance(arr, datetime.time):
            arr_min = arr.hour * 60 + arr.minute

        duration = dur if dur else (arr_min - dep_min if dep_min is not None and arr_min is not None else None)

        trips.append({
            "bus_id": str(bus_id),
            "bus_type": bus_type,
            "origin": origin,
            "dest": dest,
            "origin_loc": normalize_display_name(origin),
            "dest_loc": normalize_display_name(dest),
            "departure": dep_min,
            "arrival": arr_min,
            "duration_min": duration,
            "service": service,
            "direction": direction,
        })

    wb.close()
    return trips


def validate_trips(matrix: dict, trips: list[dict]) -> list[dict]:
    """Compare scheduled trip durations against Google Maps driving times.

    Returns list of route validations with buffer analysis.
    """
    # Group trips by unique route (origin_loc, dest_loc)
    route_trips = {}  # (origin_loc, dest_loc) -> list of trips
    for t in trips:
        if t["duration_min"] is None or t["duration_min"] <= 0:
            continue
        key = (t["origin_loc"], t["dest_loc"])
        if key not in route_trips:
            route_trips[key] = []
        route_trips[key].append(t)

    results = []
    for (o_loc, d_loc), route_list in sorted(route_trips.items()):
        gmaps = matrix.get((o_loc, d_loc))
        gmaps_min = gmaps["duration_min"] if gmaps and gmaps["duration_min"] else None

        durations = [t["duration_min"] for t in route_list]
        min_dur = min(durations)
        max_dur = max(durations)
        avg_dur = sum(durations) / len(durations)

        buffer_min = min_dur - gmaps_min if gmaps_min else None
        buffer_pct = (buffer_min / gmaps_min * 100) if gmaps_min and buffer_min is not None else None

        results.append({
            "origin": route_list[0]["origin"],
            "dest": route_list[0]["dest"],
            "origin_loc": o_loc,
            "dest_loc": d_loc,
            "num_trips": len(route_list),
            "min_scheduled": min_dur,
            "max_scheduled": max_dur,
            "avg_scheduled": round(avg_dur, 1),
            "gmaps_min": gmaps_min,
            "buffer_min": round(buffer_min, 1) if buffer_min is not None else None,
            "buffer_pct": round(buffer_pct, 1) if buffer_pct is not None else None,
            "services": list(set(t["service"] for t in route_list)),
        })

    results.sort(key=lambda r: r["buffer_min"] if r["buffer_min"] is not None else 999)
    return results


def print_validation_report(validations: list[dict]):
    """Print trip validation report to console."""
    print("\n" + "=" * 95)
    print("TRIP VALIDATION: Scheduled Duration vs Google Maps Driving Time")
    print("=" * 95)
    print(f"\n{'Route':<40} {'#':>3} {'Sched':>6} {'GMaps':>6} {'Buffer':>7} {'%':>6}  Status")
    print("-" * 95)

    critical = []
    warning = []
    ok = []

    for v in validations:
        route = f"{v['origin']} -> {v['dest']}"
        if len(route) > 39:
            route = route[:36] + "..."

        sched = f"{v['min_scheduled']:.0f}m"
        gmaps = f"{v['gmaps_min']:.0f}m" if v['gmaps_min'] else "?"
        buf = f"{v['buffer_min']:+.0f}m" if v['buffer_min'] is not None else "?"
        pct = f"{v['buffer_pct']:+.0f}%" if v['buffer_pct'] is not None else "?"

        if v['buffer_min'] is not None:
            if v['buffer_min'] < 0:
                status = "!! UNREALISTIC"
                critical.append(v)
            elif v['buffer_pct'] is not None and v['buffer_pct'] < 15:
                status = "!  TIGHT"
                warning.append(v)
            else:
                status = "OK"
                ok.append(v)
        else:
            status = "?  NO DATA"

        print(f"{route:<40} {v['num_trips']:>3} {sched:>6} {gmaps:>6} {buf:>7} {pct:>6}  {status}")

    # Summary
    print(f"\n--- Summary ---")
    print(f"Total unique routes: {len(validations)}")
    print(f"  OK (buffer >= 15%):    {len(ok)}")
    print(f"  TIGHT (buffer < 15%):  {len(warning)}")
    print(f"  UNREALISTIC (< GMaps): {len(critical)}")

    if critical:
        print(f"\n!! UNREALISTIC routes (scheduled time < Google Maps estimate):")
        for v in critical:
            print(f"   {v['origin']} -> {v['dest']}: "
                  f"scheduled {v['min_scheduled']:.0f}min vs GMaps {v['gmaps_min']:.0f}min "
                  f"({v['buffer_min']:+.0f}min / {v['buffer_pct']:+.0f}%)")
            print(f"      Services: {', '.join(v['services'])}")

    if warning:
        print(f"\n!  TIGHT routes (buffer < 15% of driving time):")
        for v in warning:
            print(f"   {v['origin']} -> {v['dest']}: "
                  f"scheduled {v['min_scheduled']:.0f}min vs GMaps {v['gmaps_min']:.0f}min "
                  f"({v['buffer_min']:+.0f}min / {v['buffer_pct']:+.0f}%)")


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

def write_excel_output(matrix: dict, locations: list[str],
                       validations: list[dict], output_file: str):
    """Write results to Excel with multiple sheets."""
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()

    # --- Sheet 1: Distance Matrix (km) ---
    ws1 = wb.active
    ws1.title = "Afstand (km)"
    _write_matrix_sheet(ws1, matrix, locations, "distance_km", "Afstand in km")

    # --- Sheet 2: Duration Matrix (min) ---
    ws2 = wb.create_sheet("Rijtijd (min)")
    _write_matrix_sheet(ws2, matrix, locations, "duration_min", "Rijtijd in minuten")

    # --- Sheet 3: Deadhead matrix ---
    ws3 = wb.create_sheet("Deadhead matrix")
    _write_deadhead_sheet(ws3, matrix, locations)

    # --- Sheet 4: Trip validation ---
    if validations:
        ws4 = wb.create_sheet("Ritvalidatie")
        _write_validation_sheet(ws4, validations)

    # --- Sheet 5: JSON data ---
    ws5 = wb.create_sheet("JSON data")
    json_data = {}
    for (o, d), vals in matrix.items():
        if o not in json_data:
            json_data[o] = {}
        json_data[o][d] = vals
    ws5.cell(1, 1, "Deadhead matrix als JSON (voor import in optimizer)")
    ws5.cell(2, 1, json.dumps(json_data, indent=2, ensure_ascii=False))

    wb.save(output_file)
    print(f"\nOutput saved to {output_file}")


def _write_matrix_sheet(ws, matrix, locations, value_key, title):
    """Write an NxN matrix sheet."""
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4",
                              fill_type="solid")
    header_font_white = Font(bold=True, size=10, color="FFFFFF")

    ws.cell(1, 1, title).font = Font(bold=True, size=12)

    for j, loc in enumerate(locations):
        cell = ws.cell(3, j + 2, loc)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", text_rotation=90)
        ws.column_dimensions[get_column_letter(j + 2)].width = 6

    for i, o_loc in enumerate(locations):
        ws.cell(i + 4, 1, o_loc).font = Font(bold=True, size=10)
        for j, d_loc in enumerate(locations):
            val = matrix.get((o_loc, d_loc), {}).get(value_key)
            if val is not None:
                ws.cell(i + 4, j + 2, val)
            ws.cell(i + 4, j + 2).alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 22


def _write_deadhead_sheet(ws, matrix, locations):
    """Write deadhead matrix in a format the optimizer can consume."""
    from openpyxl.styles import Font

    ws.cell(1, 1, "Deadhead matrix").font = Font(bold=True, size=12)
    ws.cell(2, 1, "Rijtijd in minuten tussen stations (voor lege ritten)")

    ws.cell(4, 1, "Van \\ Naar")
    for j, loc in enumerate(locations):
        ws.cell(4, j + 2, loc)

    for i, o_loc in enumerate(locations):
        ws.cell(i + 5, 1, o_loc)
        for j, d_loc in enumerate(locations):
            val = matrix.get((o_loc, d_loc), {}).get("duration_min")
            if val is not None:
                ws.cell(i + 5, j + 2, val)


def _write_validation_sheet(ws, validations):
    """Write trip validation results."""
    from openpyxl.styles import Font, Alignment, PatternFill

    ws.cell(1, 1, "Ritvalidatie: Geplande rijtijd vs Google Maps").font = Font(bold=True, size=12)

    headers = ["Van", "Naar", "Aantal ritten", "Min. gepland (min)",
               "Max. gepland (min)", "Gem. gepland (min)", "Google Maps (min)",
               "Buffer (min)", "Buffer (%)", "Status", "Busdiensten"]

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4",
                              fill_type="solid")
    header_font = Font(bold=True, size=10, color="FFFFFF")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    orange_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    for c, h in enumerate(headers, 1):
        cell = ws.cell(3, c, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for r, v in enumerate(validations, 4):
        ws.cell(r, 1, v["origin"])
        ws.cell(r, 2, v["dest"])
        ws.cell(r, 3, v["num_trips"])
        ws.cell(r, 4, v["min_scheduled"])
        ws.cell(r, 5, v["max_scheduled"])
        ws.cell(r, 6, v["avg_scheduled"])
        ws.cell(r, 7, v["gmaps_min"])
        ws.cell(r, 8, v["buffer_min"])
        ws.cell(r, 9, v["buffer_pct"])

        if v["buffer_min"] is not None:
            if v["buffer_min"] < 0:
                status = "ONREALISTISCH"
                fill = red_fill
            elif v["buffer_pct"] is not None and v["buffer_pct"] < 15:
                status = "KRAP"
                fill = orange_fill
            else:
                status = "OK"
                fill = green_fill
        else:
            status = "Geen data"
            fill = None

        ws.cell(r, 10, status)
        if fill:
            for c in range(1, 12):
                ws.cell(r, c).fill = fill

        ws.cell(r, 11, ", ".join(v["services"]))

    # Column widths
    widths = [25, 25, 12, 16, 16, 16, 16, 12, 12, 16, 40]
    for c, w in enumerate(widths, 1):
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(c)].width = w


def print_route_analysis(matrix: dict):
    """Print analysis of driving times between all station pairs."""
    print("\n" + "=" * 80)
    print("DEADHEAD DRIVING TIMES BETWEEN STATIONS")
    print("=" * 80)
    print(f"\n{'Van':<22} {'Naar':<22} {'Afstand':>8} {'Rijtijd':>8}")
    print("-" * 62)

    entries = []
    for (o, d), vals in sorted(matrix.items()):
        if o == d:
            continue
        if vals.get("duration_min") is not None:
            entries.append((o, d, vals.get("distance_km", "?"), vals["duration_min"]))

    entries.sort(key=lambda x: x[3], reverse=True)

    for o, d, km, mins in entries:
        km_str = f"{km:.1f} km" if isinstance(km, (int, float)) else "?     "
        print(f"{o:<22} {d:<22} {km_str:>8} {mins:>6.1f} min")

    print(f"\nTotal unique routes (excl. self): {len(entries)}")

    long_routes = [e for e in entries if e[3] > 30]
    if long_routes:
        print(f"\nRoutes with deadhead > 30 min ({len(long_routes)}):")
        for o, d, km, mins in long_routes:
            km_str = f"{km:.0f} km" if isinstance(km, (int, float)) else "?"
            print(f"  {o} -> {d}: {mins:.0f} min ({km_str})")


def main():
    parser = argparse.ArgumentParser(
        description="Fetch Google Maps distances between bus stations")
    parser.add_argument("--key", default=None,
                        help="Google Maps API key")
    parser.add_argument("--output", default="afstanden_stations.xlsx",
                        help="Output Excel file (default: afstanden_stations.xlsx)")
    parser.add_argument("--json-output", default="deadhead_matrix.json",
                        help="Output JSON file for deadhead matrix")
    parser.add_argument("--from-cache", default=None,
                        help="Load matrix from cached JSON instead of API call")
    parser.add_argument("--validate", default=None,
                        help="Optimizer output .xlsx file for trip validation")
    args = parser.parse_args()

    locations = sorted(STATION_ADDRESSES.keys())
    print(f"Stations: {len(locations)}")
    for loc in locations:
        print(f"  {loc}: {STATION_ADDRESSES[loc]}")

    # Load or fetch matrix
    if args.from_cache:
        print(f"\nLoading cached matrix from {args.from_cache}...")
        matrix = load_matrix_from_cache(args.from_cache)
    elif args.key:
        print(f"\nFetching {len(locations)}x{len(locations)} distance matrix "
              f"({len(locations)**2} elements)...")
        matrix = fetch_distance_matrix(args.key, locations)
        print(f"Received {len(matrix)} route entries")
        save_deadhead_json(matrix, locations, args.json_output)
    else:
        print("\nERROR: Provide --key API_KEY or --from-cache FILE")
        sys.exit(1)

    # Trip validation
    validations = []
    if args.validate:
        print(f"\nLoading trips from {args.validate}...")
        trips = load_trips_from_output(args.validate)
        print(f"Loaded {len(trips)} trips")
        if trips:
            validations = validate_trips(matrix, trips)
            print_validation_report(validations)

    # Write Excel output
    write_excel_output(matrix, locations, validations, args.output)

    # Print route analysis
    print_route_analysis(matrix)


if __name__ == "__main__":
    main()
