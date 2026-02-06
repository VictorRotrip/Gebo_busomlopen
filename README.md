# Busomloop Optimizer

Bus rotation optimizer for NS Trein Vervangend Vervoer (TVV) — replacement bus services during train disruptions.

## Overview

This tool reads a trip schedule (Bijlage J Excel format) and generates optimized bus rotations that minimize the number of buses needed while considering:

- Bus type constraints (a Dubbeldekker cannot be substituted for a Touringcar)
- Turnaround times between trips (bus-type specific)
- Reserve bus requirements per station
- Cross-service trip chaining (combining trips from different services)
- Deadhead repositioning (empty driving between stations)
- Traffic-based risk analysis (adjusting turnaround times for congestion)

The optimizer produces **5 output versions** with progressively more sophisticated optimization.

## Installation

```bash
# Required
pip install openpyxl

# Optional: for deadhead distances
pip install requests python-dotenv
```

## Quick Start

```bash
# Basic usage — generates all 5 versions
python busomloop_optimizer.py Bijlage_J.xlsx

# With custom output prefix
python busomloop_optimizer.py Bijlage_J.xlsx --output rooster_2026

# With deadhead matrix (enables version 5 cross-location optimization)
python busomloop_optimizer.py Bijlage_J.xlsx --deadhead deadhead_matrix.json

# Fast mode — only greedy algorithm for outputs 1-4
python busomloop_optimizer.py Bijlage_J.xlsx --snel
```

## Input Format

The input Excel file (Bijlage J) must have:

1. **Trip sheets** — One sheet per bus service (busdienst), containing:
   - Column headers: Busnummer, Lijn, Richting, Bustype, Snel/Stop, Patroon, etc.
   - Trip rows with departure/arrival times per station column
   - Station codes in row 3, station names in row 4, halt info in row 5

2. **Reserve sheet** — Named `Reservebussen` or `Reserve`, containing:
   - Station name, number of reserve buses, day(s), start time, end time

## Output Versions (1-5)

The optimizer generates 5 Excel files per algorithm, each building on the previous:

### Version 1: Per Dienst (Per Service)
**File:** `*_1_per_dienst.xlsx`

- Optimizes each bus service (Excel tab) separately
- No cross-service trip chaining
- Reserve buses counted as separate vehicles (not integrated)
- **Use case:** Baseline comparison, contractual per-service requirements

### Version 2: Per Dienst + Reserve Matching
**File:** `*_2_per_dienst_reservematch.xlsx`

- Same trip chaining as Version 1
- Adds **reserve-idle matching**: assigns reserve duties to buses during their idle time
- Uses bipartite matching to optimally cover reserve requirements with existing buses
- **Benefit:** Reduces total buses needed by reusing idle time for reserve standby

### Version 3: Combined + Reserves + Sensitivity
**File:** `*_3_gecombineerd_met_reserve.xlsx`

- **Cross-service optimization:** chains trips across different services
- Reserve buses planned as "phantom trips" integrated into rotations
- Includes **sensitivity analysis** sheet showing impact of turnaround time changes
- **Benefit:** Significant bus reduction through cross-service chaining

### Version 4: Combined + Risk-Based Turnaround
**File:** `*_4_gecombineerd_risico.xlsx`

- Same as Version 3, but with **traffic-aware turnaround times**
- Requires `--deadhead` with traffic data (time-slot specific travel times)
- Increases turnaround time for trips with high-traffic connections
- Includes risk analysis sheet showing per-trip risk levels (HOOG/MATIG/OK)
- **Benefit:** More robust schedules that account for traffic delays

### Version 5: Combined + Deadhead Repositioning
**File:** `*_5_gecombineerd_deadhead.xlsx`

- Enables **cross-location connections**: a bus can drive empty to another station
- Requires `--deadhead` matrix with station-to-station travel times
- Cost function penalizes deadhead time (weighted 2x vs idle time)
- Shows deadhead rows in the bus rotation schedule
- **Benefit:** Maximum flexibility, potentially fewer buses at cost of fuel/driver time

### Version 6: Fuel/Charging Constraints (NEW)
**File:** `*_6_brandstof_constraints.xlsx`

- Integrates **fuel and charging constraints** into optimization
- Tracks cumulative km per bus, validates against fuel range
- Automatically splits chains when fuel range is exceeded without refuel opportunity
- Requires `--fuel-constraints` flag and `--tanklocaties` JSON file
- ZE (Zero Emission) analysis: assigns at least 5 ZE Touringcars (NS K3 requirement)
- Calculates actual driving times to fuel/charging stations via Google Maps
- **Benefit:** Realistic fuel logistics planning, ZE feasibility assessment

## Algorithms

Two optimization algorithms are available (`--algoritme greedy|mincost|all`):

### Greedy Best-Fit (default)

```
For each trip in chronological order:
    Find the bus with the shortest idle time that can take this trip
    If found: assign trip to that bus
    Else: start a new bus
```

- **Complexity:** O(n × m) where n = trips, m = buses
- **Behavior:** Fast, deterministic, good results for most cases
- **Note:** Without deadhead, greedy and mincost produce identical results

### Min-Cost Maximum Matching (SPFA)

```
Model as bipartite graph:
    Left nodes = trips (as "trip ends")
    Right nodes = trips (as "trip starts")
    Edge (i,j) exists if trip i can connect to trip j
    Edge cost = deadhead_time × 2 + idle_time

Find maximum matching with minimum total cost using:
    Successive Shortest Path Algorithm (SPFA)
```

- **Complexity:** O(n² × m) worst case, typically much faster
- **Behavior:** Optimal for minimizing buses first, then total idle+deadhead time
- **Note:** Essential when deadhead repositioning is enabled

## Key Assumptions

### Turnaround Times

Minimum time required between a bus arriving and departing for the next trip:

| Bus Type | Default (min) | Description |
|----------|--------------|-------------|
| Dubbeldekker | 8 | Double-decker bus, large, needs more time |
| Touringcar | 6 | Coach, medium turnaround |
| Lagevloerbus | 5 | Low-floor / articulated bus |
| Midi bus | 4 | Medium-sized bus |
| Taxibus | 3 | Small bus, quick turnaround |

- **Auto-detection:** If not specified, turnaround times are detected from the input data (smallest gap between arrival and departure at the same location)
- **Absolute minimum:** 2 minutes for any real trip connection
- **Override:** Use `--keer-dd`, `--keer-tc`, `--keer-lvb`, `--keer-midi`, `--keer-taxi`

### Connection Feasibility (`can_connect`)

A bus can connect trip A to trip B if ALL of the following are true:

1. **Same bus type:** A.bus_type == B.bus_type
2. **Same date:** A.date == B.date
3. **Time feasible:** B.departure >= A.arrival + turnaround + deadhead_time
4. **Location feasible:**
   - Same location: A.destination == B.origin (no deadhead needed)
   - Different location: deadhead matrix must exist and provide travel time

### Reserve Bus Matching

Reserve requirements specify: station, count, day(s), start time, end time.

**Matching logic:**
- A bus can cover a reserve requirement if it has an idle window overlapping the reserve period at that station
- Uses bipartite matching to maximize coverage
- Uncovered reserves require additional standalone buses

### Cost Function (Min-Cost Algorithm)

```python
cost = deadhead_time × 2 + idle_time
```

- **Deadhead penalty (2x):** Empty driving costs fuel and driver time
- **Idle time:** Waiting time between trips (driver still on duty)
- **Primary objective:** Minimize number of buses (maximize matching)
- **Secondary objective:** Minimize total cost (break ties)

## Output Excel Structure

Each output file contains these sheets:

| Sheet | Description |
|-------|-------------|
| **Busomloop** | Per-bus rotation schedule (Transvision style) |
| **Ritsamenhang** | Trip connections showing which trips chain together |
| **Berekeningen** | Summary statistics, algorithm details, turnaround analysis |
| **Sensitiviteit** | (V3+) Impact of turnaround changes on bus count |
| **Risico-analyse** | (V4+) Per-trip traffic risk and turnaround adjustments |

## Additional Tools

### Google Maps Distance Fetcher

Fetches station-to-station travel times and distances. This data is used for:
1. **Deadhead repositioning** — empty drives between stations
2. **Trip validation / risk analysis** — comparing scheduled NS trip times vs actual Google Maps times
3. **Fuel calculations** — estimating km driven using actual distances

```bash
# Basic matrix: baseline times + distances (uses GOOGLE_MAPS_API_KEY from .env)
python google_maps_distances.py --input Bijlage_J.xlsx
# Output: deadhead_matrix.json

# Traffic-aware matrix: times per time slot + distances (recommended)
python google_maps_distances.py --input Bijlage_J.xlsx --traffic
# Output: traffic_matrix.json
```

**Difference between the two:**

| | `deadhead_matrix.json` | `traffic_matrix.json` |
|---|---|---|
| **Travel times** | Baseline only (no traffic) | 6 time slots + baseline |
| **Distances (km)** | ✅ Yes | ✅ Yes |
| **Risk analysis** | ❌ No (no time slots) | ✅ Yes |
| **Fuel estimates** | ✅ Yes | ✅ Yes |

**Time slots in traffic matrix:**
- `nacht` (00:00-06:00), `ochtendspits` (07:00-09:00), `dal` (10:00-15:00)
- `middagspits` (16:00-18:00), `avond` (19:00-23:00), `weekend`

The **risk analysis** (Version 4+) compares each scheduled NS trip against the appropriate time slot. For example, a trip departing at 08:15 is compared against "ochtendspits" times — if the scheduled duration is shorter than Google Maps predicts with traffic, it's flagged as high risk.

### Financial Input Configuration

All financial and operational variables are stored in `additional_inputs.xlsx` (5 sheets):

| Sheet | Contents |
|-------|----------|
| **Tarieven** | Hourly rates per bus type (from Prijzenblad) |
| **Chauffeurkosten** | CAO variables (pauzestaffel, ORT, overtime) |
| **Buskosten** | Fuel consumption, tank capacity, ZE range, bus speed factors |
| **Duurzaamheid** | ZE/HVO incentives, KPI targets, malus rules |
| **Brandstofprijzen** | Current fuel/electricity prices (auto-update available) |

```bash
# Auto-update fuel/electricity prices from APIs
python fetch_fuel_charging_prices.py
```

### Fuel/Charging Station Fetcher

Fetches nearby fuel stations and EV chargers for each bus station:

```bash
# Auto-discover stations from input Excel
python fetch_tanklocaties.py --input Bijlage_J.xlsx

# With Google Maps driving distances (recommended)
python fetch_tanklocaties.py --input Bijlage_J.xlsx --gmaps

# With specific radius
python fetch_tanklocaties.py --input Bijlage_J.xlsx --radius 10 --gmaps

# Output: tanklocaties.json
```

When using `--gmaps`, each fuel/charging station includes:
- `drive_time_min`: actual driving time from bus station
- `drive_distance_km`: actual driving distance
- These are used for accurate refuel/charge feasibility calculations

## Future Versions (7-9)

Version 6 (fuel/charging constraints) is implemented. See `PLAN_FINANCIAL_OPTIMIZATION.md` for planned financial optimization:

| Version | Status | Description |
|---------|--------|-------------|
| 6 | ✅ DONE | Fuel/charging constraints, ZE feasibility, Google Maps integration |
| 7 | Planned | Financial analysis overlay (revenue, costs, profit calculation) |
| 8 | Planned | Euro-based cost optimization (driver costs, fuel costs, ORT surcharges) |
| 9 | Planned | Full profit optimization (may use different bus count for profit) |

## Command Line Reference

```
usage: busomloop_optimizer.py [-h] [--output OUTPUT] [--algoritme {greedy,mincost,all}]
                               [--deadhead DEADHEAD] [--traffic-matrix TRAFFIC]
                               [--tanklocaties JSON] [--inputs XLSX]
                               [--ze] [--fuel-constraints] [--snel]
                               [--keer-dd MIN] [--keer-tc MIN] [--keer-lvb MIN]
                               [--keer-midi MIN] [--keer-taxi MIN]
                               invoer.xlsx

Arguments:
  invoer.xlsx              Input Excel file (Bijlage J format)

Options:
  --output, -o OUTPUT      Output file prefix (default: busomloop_output)
  --algoritme {greedy,mincost,all}
                           Optimization algorithm (default: greedy)
  --deadhead DEADHEAD      JSON file with deadhead travel times
  --traffic-matrix TRAFFIC JSON file with traffic-aware travel times (includes distances)
  --tanklocaties JSON      JSON file with fuel/charging station locations
  --inputs XLSX            additional_inputs.xlsx with financial/operational config
  --ze                     Enable ZE (Zero Emission) feasibility analysis
  --fuel-constraints       Enable diesel fuel range validation
  --snel                   Fast mode: skip non-greedy for versions 1-4
  --keer-dd MIN            Turnaround time Dubbeldekker (default: 8)
  --keer-tc MIN            Turnaround time Touringcar (default: 6)
  --keer-lvb MIN           Turnaround time Lagevloerbus (default: 5)
  --keer-midi MIN          Turnaround time Midi bus (default: 4)
  --keer-taxi MIN          Turnaround time Taxibus (default: 3)
```

## Example Workflow

```bash
# 1. Fetch station-to-station distances with traffic data (GOOGLE_MAPS_API_KEY in .env)
python google_maps_distances.py --input Bijlage_J.xlsx --traffic

# 2. Fetch fuel/charging stations with driving times
python fetch_tanklocaties.py --input Bijlage_J.xlsx --gmaps

# 3. Run optimizer with all features (versions 1-6)
python busomloop_optimizer.py Bijlage_J.xlsx \
    --deadhead deadhead_matrix.json \
    --traffic-matrix traffic_matrix.json \
    --tanklocaties tanklocaties.json \
    --fuel-constraints \
    --ze \
    --output project_x

# 4. Compare results
# - project_x_greedy_1_per_dienst.xlsx      → baseline per service
# - project_x_greedy_3_gecombineerd.xlsx    → cross-service optimization
# - project_x_greedy_5_deadhead.xlsx        → with deadhead repositioning
# - project_x_greedy_*_ze.xlsx              → with ZE feasibility analysis
# - project_x_greedy_*_fuel.xlsx            → with fuel constraint validation
```

### Data Flow (Version 6)

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           Google Maps API                                   │
└─────────────────────────┬───────────────────────┬───────────────────────────┘
                          │                       │
                          ▼                       ▼
              ┌───────────────────────┐   ┌───────────────────────┐
              │ traffic_matrix.json   │   │ tanklocaties.json     │
              │ - time slots (min)    │   │ - fuel stations       │
              │ - distances_km        │   │ - charging stations   │
              │ - baseline times      │   │ - drive_time_min      │
              └───────────┬───────────┘   │ - drive_distance_km   │
                          │               └───────────┬───────────┘
                          │                           │
                          ▼                           ▼
              ┌───────────────────────────────────────────────────┐
              │           busomloop_optimizer.py                  │
              │  - Calculate avg_speed from Google Maps data      │
              │  - Apply bus speed factors (from Excel)           │
              │  - Validate fuel range using actual distances     │
              │  - Check refuel feasibility with drive times      │
              │  - Assign ZE buses with charging strategy         │
              └───────────────────────────────────────────────────┘
```

## License

Internal use only — Gebo project for NS TVV optimization.
