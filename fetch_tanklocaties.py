#!/usr/bin/env python3
"""
fetch_tanklocaties.py

Fetches fuel station (tankstation) and EV charging station (laadpaal) locations
near the bus stations used in the roster. Outputs a JSON file that versions 7-9
of the optimizer can use for fueling/charging logistics planning.

Data sources:
  - OpenStreetMap Overpass API: fuel stations (diesel, HVO100, LPG, etc.) — free, no auth
  - Open Charge Map API: EV charging stations with power/connector info — free, optional API key
  - Nominatim (OSM): geocoding bus station names to lat/lon — free, no auth
  - Google Maps Distance Matrix API: actual driving distances/times (optional, requires API key)

Usage:
  # Auto-discover stations from input Excel + fetch nearby fuel/charging:
  python fetch_tanklocaties.py --input Bijlage_J.xlsx

  # Use a JSON file with station coordinates:
  python fetch_tanklocaties.py --coords station_coords.json

  # Specify stations manually:
  python fetch_tanklocaties.py --stations "Utrecht Centraal" "Ede-Wageningen" "Amersfoort"

  # Adjust search radius (default 5 km):
  python fetch_tanklocaties.py --input Bijlage_J.xlsx --radius 10

  # With Open Charge Map API key (optional, for higher rate limit):
  python fetch_tanklocaties.py --input Bijlage_J.xlsx --ocm-key YOUR_KEY

  # With Google Maps driving distances (uses GOOGLE_MAPS_API_KEY from .env):
  python fetch_tanklocaties.py --input Bijlage_J.xlsx --gmaps

Output:
  tanklocaties.json — structured JSON with nearby fuel and charging stations per bus station
  When --gmaps is used, each station includes drive_distance_km and drive_time_min fields

Requirements:
  pip install requests openpyxl
"""

import argparse
import json
import math
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

try:
    import requests
except ImportError:
    sys.exit("Error: requests not installed. Run: pip install requests")


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DEFAULT_RADIUS_KM = 5
DEFAULT_OUTPUT = "tanklocaties.json"

# Overpass API for OpenStreetMap queries
OVERPASS_URL = "https://overpass-api.de/api/interpreter"

# Open Charge Map API
OCM_URL = "https://api.openchargemap.io/v3/poi/"

# Nominatim (OSM geocoding) — respect usage policy: max 1 req/sec, User-Agent required
NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"

# Netherlands bounding box for validation
NL_BOUNDS = {"lat_min": 50.5, "lat_max": 53.7, "lon_min": 3.3, "lon_max": 7.3}

# Google Maps Distance Matrix API
GMAPS_DISTANCE_URL = "https://maps.googleapis.com/maps/api/distancematrix/json"


def load_dotenv():
    """Load .env file from the script's directory into os.environ."""
    import os
    env_path = Path(__file__).parent / ".env"
    if not env_path.exists():
        return
    with open(env_path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                key, _, value = line.partition("=")
                key = key.strip()
                value = value.strip().strip("'\"")
                os.environ.setdefault(key, value)

# User-Agent for Nominatim (required by their usage policy)
USER_AGENT = "BusOmloopOptimizer/1.0 (fuel-station-fetcher; contact: busomloop@gebo.nl)"


# ---------------------------------------------------------------------------
# Geocoding: station names → lat/lon
# ---------------------------------------------------------------------------

def geocode_station_nominatim(station_name: str) -> Optional[dict]:
    """Geocode a Dutch bus station name using OSM Nominatim.

    Returns {"lat": float, "lon": float} or None if not found.
    """
    # Try with "station" prefix for better results on transit stops
    queries = [
        f"Station {station_name}, Nederland",
        f"{station_name} busstation, Nederland",
        f"{station_name}, Nederland",
    ]

    for query in queries:
        try:
            resp = requests.get(
                NOMINATIM_URL,
                params={
                    "q": query,
                    "format": "json",
                    "countrycodes": "nl",
                    "limit": 1,
                    "addressdetails": 0,
                },
                headers={"User-Agent": USER_AGENT},
                timeout=15,
            )
            resp.raise_for_status()
            results = resp.json()
        except requests.RequestException:
            continue

        if results:
            lat = float(results[0]["lat"])
            lon = float(results[0]["lon"])
            # Validate within NL bounds
            if (NL_BOUNDS["lat_min"] <= lat <= NL_BOUNDS["lat_max"]
                    and NL_BOUNDS["lon_min"] <= lon <= NL_BOUNDS["lon_max"]):
                return {"lat": lat, "lon": lon}

        # Rate limit: Nominatim requires max 1 request/sec
        time.sleep(1.1)

    return None


def geocode_stations(station_names: List[str]) -> dict:
    """Geocode a list of station names. Returns {name: {"lat": .., "lon": ..}}."""
    print(f"\nGeocodering van {len(station_names)} stations via Nominatim...")
    results = {}
    for name in station_names:
        print(f"  {name}...", end=" ", flush=True)
        coords = geocode_station_nominatim(name)
        if coords:
            print(f"OK ({coords['lat']:.4f}, {coords['lon']:.4f})")
            results[name] = coords
        else:
            print("NIET GEVONDEN")
        time.sleep(0.5)  # Extra politeness delay
    print(f"  {len(results)}/{len(station_names)} stations geocodeerd.")
    return results


def load_coords_from_json(path: str) -> dict:
    """Load station coordinates from a JSON file.

    Expected format:
    {
      "Utrecht Centraal": {"lat": 52.0907, "lon": 5.1214},
      "Ede-Wageningen": {"lat": 52.0384, "lon": 5.6522},
      ...
    }
    """
    with open(path, "r") as f:
        data = json.load(f)
    print(f"  {len(data)} stations geladen uit {path}")
    return data


# ---------------------------------------------------------------------------
# Haversine distance
# ---------------------------------------------------------------------------

def haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Calculate distance between two points in km using the haversine formula."""
    R = 6371.0  # Earth radius in km
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat / 2) ** 2
         + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2))
         * math.sin(dlon / 2) ** 2)
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


# ---------------------------------------------------------------------------
# OpenStreetMap Overpass: Fuel stations
# ---------------------------------------------------------------------------

def fetch_fuel_stations_osm(lat: float, lon: float, radius_m: int = 5000,
                            max_retries: int = 3) -> List[dict]:
    """Fetch fuel stations near a point using OSM Overpass API.

    Returns list of dicts with station info including fuel types.
    Includes retry logic with exponential backoff for rate limiting.
    """
    # Overpass QL: find fuel amenities within radius
    # We query both nodes and ways (some stations are mapped as areas)
    query = f"""
    [out:json][timeout:30];
    (
      node["amenity"="fuel"](around:{radius_m},{lat},{lon});
      way["amenity"="fuel"](around:{radius_m},{lat},{lon});
    );
    out center tags;
    """

    data = None
    for attempt in range(max_retries):
        try:
            resp = requests.post(
                OVERPASS_URL,
                data={"data": query},
                timeout=45,
                headers={"User-Agent": USER_AGENT},
            )
            resp.raise_for_status()
            data = resp.json()
            break  # Success
        except requests.RequestException as e:
            wait_time = 2 ** (attempt + 1)  # 2, 4, 8 seconds
            if attempt < max_retries - 1:
                status_code = getattr(e.response, 'status_code', None) if hasattr(e, 'response') else None
                if status_code in (429, 504, 503):
                    print(f"    [OSM] Rate limited, wacht {wait_time}s en probeer opnieuw...")
                    time.sleep(wait_time)
                    continue
            print(f"    [OSM] Overpass query failed: {e}")
            return []

    if data is None:
        return []

    stations = []
    for element in data.get("elements", []):
        tags = element.get("tags", {})

        # Get coordinates (nodes have lat/lon directly, ways have center)
        if element["type"] == "node":
            s_lat = element["lat"]
            s_lon = element["lon"]
        elif "center" in element:
            s_lat = element["center"]["lat"]
            s_lon = element["center"]["lon"]
        else:
            continue

        # Extract fuel types from OSM tags
        fuels = {}
        fuel_tags = {
            "diesel": ["fuel:diesel", "fuel:Diesel"],
            "hvo100": ["fuel:HVO100", "fuel:hvo100", "fuel:GTL_diesel", "fuel:gtl"],
            "lpg": ["fuel:lpg", "fuel:LPG"],
            "adblue": ["fuel:adblue", "fuel:AdBlue"],
            "cng": ["fuel:cng", "fuel:CNG"],
            "lng": ["fuel:lng", "fuel:LNG"],
            "e10": ["fuel:e10", "fuel:octane_95"],
            "e5": ["fuel:e5", "fuel:octane_98"],
        }

        for fuel_name, tag_keys in fuel_tags.items():
            for tag_key in tag_keys:
                if tags.get(tag_key) == "yes":
                    fuels[fuel_name] = True
                    break

        # If no specific fuel tags, assume at least diesel for truck stops
        # (most Dutch fuel stations have diesel but not all tag it)
        if not fuels and tags.get("amenity") == "fuel":
            fuels["diesel"] = True  # Assumed — not tagged explicitly

        distance = haversine_km(lat, lon, s_lat, s_lon)

        station_info = {
            "osm_id": element.get("id"),
            "name": tags.get("name", "Onbekend tankstation"),
            "brand": tags.get("brand", tags.get("operator", "")),
            "lat": round(s_lat, 6),
            "lon": round(s_lon, 6),
            "distance_km": round(distance, 2),
            "fuels": fuels,
            "has_hvo100": fuels.get("hvo100", False),
            "has_diesel": fuels.get("diesel", False),
            "opening_hours": tags.get("opening_hours", ""),
            "address": _build_address_from_tags(tags),
        }
        stations.append(station_info)

    # Sort by distance
    stations.sort(key=lambda s: s["distance_km"])
    return stations


def _build_address_from_tags(tags: dict) -> str:
    """Build a readable address string from OSM tags."""
    parts = []
    street = tags.get("addr:street", "")
    housenumber = tags.get("addr:housenumber", "")
    if street:
        parts.append(f"{street} {housenumber}".strip())
    city = tags.get("addr:city", tags.get("addr:place", ""))
    if city:
        parts.append(city)
    postcode = tags.get("addr:postcode", "")
    if postcode:
        parts.append(postcode)
    return ", ".join(parts) if parts else ""


# ---------------------------------------------------------------------------
# Open Charge Map: EV Charging stations
# ---------------------------------------------------------------------------

def fetch_charging_stations_ocm(lat: float, lon: float, radius_km: float = 5,
                                 api_key: str = None) -> List[dict]:
    """Fetch EV charging stations near a point using Open Charge Map API.

    Returns list of dicts with charger info including power and connectors.
    """
    params = {
        "output": "json",
        "latitude": lat,
        "longitude": lon,
        "distance": radius_km,
        "distanceunit": "KM",
        "maxresults": 100,
        "countrycode": "NL",
        "compact": "true",
        "verbose": "false",
    }
    if api_key:
        params["key"] = api_key

    headers = {"User-Agent": USER_AGENT}

    try:
        resp = requests.get(OCM_URL, params=params, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
    except requests.RequestException as e:
        print(f"    [OCM] Open Charge Map query failed: {e}")
        return []

    stations = []
    for poi in data:
        addr_info = poi.get("AddressInfo", {})
        s_lat = addr_info.get("Latitude")
        s_lon = addr_info.get("Longitude")
        if s_lat is None or s_lon is None:
            continue

        # Extract connector details
        connections = poi.get("Connections", [])
        connectors = []
        max_power_kw = 0
        num_points = 0

        for conn in connections:
            conn_type = conn.get("ConnectionType", {})
            type_title = conn_type.get("Title", "Unknown")
            power = conn.get("PowerKW") or 0
            quantity = conn.get("Quantity") or 1

            connectors.append({
                "type": type_title,
                "power_kw": power,
                "quantity": quantity,
            })
            max_power_kw = max(max_power_kw, power)
            num_points += quantity

        # Determine charger category
        if max_power_kw >= 150:
            category = "ultra_fast"   # ≥150 kW (HPC)
        elif max_power_kw >= 50:
            category = "fast"         # 50-149 kW (DC fast)
        elif max_power_kw >= 22:
            category = "semi_fast"    # 22-49 kW (AC fast / DC)
        else:
            category = "slow"         # <22 kW (AC normal)

        distance = haversine_km(lat, lon, s_lat, s_lon)

        # Extract operator name
        operator_info = poi.get("OperatorInfo", {})
        operator_name = operator_info.get("Title", "") if operator_info else ""

        station_info = {
            "ocm_id": poi.get("ID"),
            "name": addr_info.get("Title", "Onbekend laadpunt"),
            "operator": operator_name,
            "lat": round(s_lat, 6),
            "lon": round(s_lon, 6),
            "distance_km": round(distance, 2),
            "max_power_kw": max_power_kw,
            "category": category,
            "num_points": num_points,
            "connectors": connectors,
            "address": addr_info.get("AddressLine1", ""),
            "town": addr_info.get("Town", ""),
            "is_operational": poi.get("StatusType", {}).get("IsOperational", True)
                             if poi.get("StatusType") else True,
        }
        stations.append(station_info)

    # Sort by distance
    stations.sort(key=lambda s: s["distance_km"])
    return stations


# ---------------------------------------------------------------------------
# Google Maps driving distances
# ---------------------------------------------------------------------------

def fetch_gmaps_distances(origin_lat: float, origin_lon: float,
                          stations: list, gmaps_key: str,
                          max_destinations: int = 25) -> list:
    """Fetch actual driving distances from origin to fuel/charging stations.

    Uses Google Maps Distance Matrix API to get accurate driving distance and
    duration, replacing the straight-line distance estimates.

    Args:
        origin_lat, origin_lon: Bus station coordinates
        stations: List of fuel/charging station dicts with lat/lon
        gmaps_key: Google Maps API key
        max_destinations: Max destinations per API call (API limit is 25)

    Returns:
        Updated list of stations with drive_distance_km and drive_time_min added
    """
    if not stations or not gmaps_key:
        return stations

    # Process in batches (API limit is 25 destinations per request)
    origin = f"{origin_lat},{origin_lon}"

    for i in range(0, len(stations), max_destinations):
        batch = stations[i:i + max_destinations]
        destinations = "|".join(f"{s['lat']},{s['lon']}" for s in batch)

        try:
            resp = requests.get(
                GMAPS_DISTANCE_URL,
                params={
                    "origins": origin,
                    "destinations": destinations,
                    "key": gmaps_key,
                    "mode": "driving",
                    "language": "nl",
                },
                timeout=30
            )
            resp.raise_for_status()
            data = resp.json()

            if data["status"] != "OK":
                print(f"      Google Maps API fout: {data['status']}")
                continue

            # Parse results
            for j, element in enumerate(data["rows"][0]["elements"]):
                if element["status"] == "OK":
                    batch[j]["drive_distance_km"] = round(
                        element["distance"]["value"] / 1000, 2
                    )
                    batch[j]["drive_time_min"] = round(
                        element["duration"]["value"] / 60, 1
                    )

            time.sleep(0.2)  # Rate limit

        except requests.RequestException as e:
            print(f"      Google Maps verzoek mislukt: {e}")
            continue

    return stations


# ---------------------------------------------------------------------------
# Bulk fetch: all stations for all bus stops
# ---------------------------------------------------------------------------

def fetch_all_nearby(station_coords: dict, radius_km: float = 5,
                     ocm_key: str = None,
                     gmaps_key: str = None,
                     fuel_only: bool = False,
                     charging_only: bool = False,
                     min_charger_power: int = 0) -> dict:
    """Fetch fuel + charging stations near all bus stations.

    Args:
        station_coords: {station_name: {"lat": .., "lon": ..}}
        radius_km: search radius in km
        ocm_key: optional Open Charge Map API key
        gmaps_key: optional Google Maps API key for driving distances
        fuel_only: only fetch fuel stations
        charging_only: only fetch charging stations
        min_charger_power: minimum charger power in kW (0 = no filter)

    Returns: {station_name: {"lat": .., "lon": .., "fuel_stations": [...], "charging_stations": [...]}}
    """
    radius_m = int(radius_km * 1000)
    results = {}
    total = len(station_coords)

    for i, (name, coords) in enumerate(station_coords.items(), 1):
        lat, lon = coords["lat"], coords["lon"]
        print(f"\n[{i}/{total}] {name} ({lat:.4f}, {lon:.4f})")

        entry = {
            "lat": lat,
            "lon": lon,
            "fuel_stations": [],
            "charging_stations": [],
        }

        # Fetch fuel stations from OSM
        if not charging_only:
            print(f"  Tankstations ophalen (OSM, radius {radius_km} km)...")
            fuel = fetch_fuel_stations_osm(lat, lon, radius_m)
            entry["fuel_stations"] = fuel
            n_hvo = sum(1 for s in fuel if s.get("has_hvo100"))
            print(f"    {len(fuel)} tankstations gevonden ({n_hvo} met HVO100)")
            time.sleep(2)  # Rate limit for Overpass (increased to avoid 429 errors)

            # Fetch Google Maps driving distances if API key provided
            if gmaps_key and fuel:
                print(f"    Rijtijden ophalen via Google Maps...")
                entry["fuel_stations"] = fetch_gmaps_distances(lat, lon, fuel, gmaps_key)
                n_with_drive = sum(1 for s in entry["fuel_stations"] if s.get("drive_time_min"))
                print(f"      {n_with_drive}/{len(fuel)} met rijtijd")

        # Fetch charging stations from Open Charge Map
        if not fuel_only:
            power_filter = f" ≥{min_charger_power} kW" if min_charger_power > 0 else ""
            print(f"  Laadstations ophalen (Open Charge Map, radius {radius_km} km{power_filter})...")
            charging = fetch_charging_stations_ocm(lat, lon, radius_km, api_key=ocm_key)

            # Filter by minimum power if specified
            if min_charger_power > 0:
                charging_before = len(charging)
                charging = [s for s in charging if s.get("max_power_kw", 0) >= min_charger_power]
                print(f"    {charging_before} gevonden, {len(charging)} met ≥{min_charger_power} kW")
            else:
                n_fast = sum(1 for s in charging if s.get("max_power_kw", 0) >= 50)
                print(f"    {len(charging)} laadstations gevonden ({n_fast} snelladers ≥50 kW)")

            entry["charging_stations"] = charging
            time.sleep(0.5)

            # Fetch Google Maps driving distances if API key provided
            if gmaps_key and charging:
                print(f"    Rijtijden ophalen via Google Maps...")
                entry["charging_stations"] = fetch_gmaps_distances(lat, lon, charging, gmaps_key)
                n_with_drive = sum(1 for s in entry["charging_stations"] if s.get("drive_time_min"))
                print(f"      {n_with_drive}/{len(charging)} met rijtijd")

        results[name] = entry

    return results


# ---------------------------------------------------------------------------
# Summary & statistics
# ---------------------------------------------------------------------------

def print_summary(results: dict) -> None:
    """Print a human-readable summary of fetched station data."""
    print(f"\n{'='*80}")
    print("SAMENVATTING TANKLOCATIES & LAADSTATIONS")
    print(f"{'='*80}")

    total_fuel = 0
    total_hvo = 0
    total_charging = 0
    total_fast = 0

    for name, data in sorted(results.items()):
        fuel = data.get("fuel_stations", [])
        charging = data.get("charging_stations", [])
        n_hvo = sum(1 for s in fuel if s.get("has_hvo100"))
        n_fast = sum(1 for s in charging if s.get("max_power_kw", 0) >= 50)

        total_fuel += len(fuel)
        total_hvo += n_hvo
        total_charging += len(charging)
        total_fast += n_fast

        print(f"\n  {name}:")
        print(f"    Tankstations:  {len(fuel):>3}  (waarvan {n_hvo} met HVO100)")
        print(f"    Laadstations:  {len(charging):>3}  (waarvan {n_fast} snelladers ≥50 kW)")

        # Show closest fuel station
        if fuel:
            closest = fuel[0]
            print(f"    Dichtstbijzijnde tank: {closest['name']} "
                  f"({closest['brand']}) — {closest['distance_km']} km"
                  + (" [HVO100]" if closest.get("has_hvo100") else ""))

        # Show closest HVO100 station
        hvo_stations = [s for s in fuel if s.get("has_hvo100")]
        if hvo_stations:
            closest_hvo = hvo_stations[0]
            print(f"    Dichtstbijzijnde HVO100: {closest_hvo['name']} "
                  f"({closest_hvo['brand']}) — {closest_hvo['distance_km']} km")

        # Show closest fast charger
        fast_chargers = [s for s in charging if s.get("max_power_kw", 0) >= 50]
        if fast_chargers:
            closest_fc = fast_chargers[0]
            print(f"    Dichtstbijzijnde snellader: {closest_fc['name']} "
                  f"({closest_fc['operator']}) — {closest_fc['distance_km']} km, "
                  f"{closest_fc['max_power_kw']} kW")

    print(f"\n{'—'*80}")
    print(f"TOTAAL: {total_fuel} tankstations ({total_hvo} HVO100), "
          f"{total_charging} laadstations ({total_fast} snelladers)")
    print(f"{'—'*80}")


# ---------------------------------------------------------------------------
# Extract stations from input Excel (reuse optimizer's parser)
# ---------------------------------------------------------------------------

def extract_stations_from_excel(input_file: str) -> List[str]:
    """Extract station display names from the input Excel using the optimizer's parser."""
    try:
        from busomloop_optimizer import parse_all_sheets, build_station_registry
    except ImportError:
        sys.exit("Error: cannot import busomloop_optimizer. "
                 "Make sure busomloop_optimizer.py is in the same directory.")

    print(f"Stations extraheren uit {input_file}...")
    all_trips, reserves, _ = parse_all_sheets(input_file)
    registry = build_station_registry(all_trips, reserves)
    names = sorted(registry.values())
    print(f"  {len(names)} stations gevonden: {', '.join(names)}")
    return names


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Fetch nearby fuel and charging station locations for bus roster stations. "
                    "Uses OpenStreetMap (fuel) and Open Charge Map (EV charging)."
    )

    # Station source (pick one)
    source_group = parser.add_mutually_exclusive_group(required=True)
    source_group.add_argument(
        "--input", "-i",
        help="Input Excel file (Bijlage J) — auto-discovers stations from roster"
    )
    source_group.add_argument(
        "--coords",
        help="JSON file with station coordinates: {name: {lat, lon}}"
    )
    source_group.add_argument(
        "--stations", nargs="+",
        help="Station names to look up (space-separated)"
    )

    # Options
    parser.add_argument("--radius", type=float, default=DEFAULT_RADIUS_KM,
                        help=f"Search radius in km (default: {DEFAULT_RADIUS_KM})")
    parser.add_argument("--ocm-key", default=None,
                        help="Open Charge Map API key (optional, for higher rate limit)")
    parser.add_argument("--gmaps", action="store_true",
                        help="Fetch actual driving distances via Google Maps API "
                             "(uses GOOGLE_MAPS_API_KEY from .env)")
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT,
                        help=f"Output JSON file (default: {DEFAULT_OUTPUT})")
    parser.add_argument("--fuel-only", action="store_true",
                        help="Only fetch fuel stations (skip charging)")
    parser.add_argument("--min-charger-power", type=int, default=0,
                        help="Minimum charger power in kW (e.g., 150 for ultra-fast only)")
    parser.add_argument("--charging-only", action="store_true",
                        help="Only fetch charging stations (skip fuel)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Geocode stations but don't fetch fuel/charging data")

    args = parser.parse_args()

    # Load API keys from .env file
    import os
    load_dotenv()

    # Google Maps API key
    gmaps_key = None
    if args.gmaps:
        gmaps_key = os.environ.get("GOOGLE_MAPS_API_KEY")
        if not gmaps_key:
            print("WAARSCHUWING: --gmaps opgegeven maar GOOGLE_MAPS_API_KEY niet gevonden in .env")
            print("  Rijtijden worden niet opgehaald.")

    # Open Charge Map API key (from --ocm-key flag or OCM_API_KEY in .env)
    ocm_key = args.ocm_key or os.environ.get("OCM_API_KEY")

    print("=" * 70)
    print("Tanklocaties & Laadstations Fetcher")
    print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Zoekradius: {args.radius} km")
    if args.min_charger_power > 0:
        print(f"Minimum lader vermogen: {args.min_charger_power} kW (ultra-fast)")
    if gmaps_key:
        print("Google Maps: rijtijden worden opgehaald")
    if ocm_key:
        print("Open Charge Map: API key geconfigureerd")
    print("=" * 70)

    # Step 1: Get station coordinates
    if args.coords:
        print(f"\nCoordinaten laden uit {args.coords}...")
        station_coords = load_coords_from_json(args.coords)

    elif args.input:
        station_names = extract_stations_from_excel(args.input)
        station_coords = geocode_stations(station_names)
        if not station_coords:
            sys.exit("Error: geen stations konden worden geocodeerd. "
                     "Gebruik --coords met een JSON bestand met coordinaten.")

    else:
        station_coords = geocode_stations(args.stations)
        if not station_coords:
            sys.exit("Error: geen stations konden worden geocodeerd.")

    print(f"\n{len(station_coords)} stations met coordinaten beschikbaar.")

    if args.dry_run:
        print("\n[DRY RUN] Geocodering compleet. Geen fuel/charging data opgehaald.")
        # Still save the coordinates
        output = {
            "metadata": {
                "fetched_at": datetime.now().isoformat(),
                "radius_km": args.radius,
                "sources": [],
                "dry_run": True,
            },
            "stations": {
                name: {"lat": c["lat"], "lon": c["lon"],
                       "fuel_stations": [], "charging_stations": []}
                for name, c in station_coords.items()
            },
        }
        with open(args.output, "w") as f:
            json.dump(output, f, indent=2, ensure_ascii=False)
        print(f"Coordinaten opgeslagen in {args.output}")
        return

    # Step 2: Fetch nearby fuel and charging stations
    results = fetch_all_nearby(
        station_coords,
        radius_km=args.radius,
        ocm_key=ocm_key,
        gmaps_key=gmaps_key,
        fuel_only=args.fuel_only,
        charging_only=args.charging_only,
        min_charger_power=args.min_charger_power,
    )

    # Step 3: Print summary
    print_summary(results)

    # Step 4: Save to JSON
    sources = []
    if not args.charging_only:
        sources.append("OpenStreetMap Overpass")
    if not args.fuel_only:
        sources.append("Open Charge Map")
    if gmaps_key:
        sources.append("Google Maps Distance Matrix")

    output = {
        "metadata": {
            "fetched_at": datetime.now().isoformat(),
            "radius_km": args.radius,
            "min_charger_power_kw": args.min_charger_power,
            "num_bus_stations": len(results),
            "sources": sources,
            "has_drive_times": gmaps_key is not None,
        },
        "stations": results,
    }

    output_path = Path(args.output)
    with open(output_path, "w") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    print(f"\nResultaten opgeslagen in {output_path}")
    print(f"Gebruik dit bestand als input voor de optimizer (versie 8-9).")


if __name__ == "__main__":
    main()
