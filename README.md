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

Fetches station-to-station travel times for the deadhead matrix:

```bash
# Verify addresses first
python google_maps_distances.py --input Bijlage_J.xlsx --key YOUR_API_KEY --verify

# Fetch full distance matrix
python google_maps_distances.py --input Bijlage_J.xlsx --key YOUR_API_KEY

# Output: deadhead_matrix.json, afstanden_stations.xlsx
```

### Financial Input Generator (Versions 7-9 preparation)

Creates a focused Excel with essential financial variables:

```bash
# Generate financieel_input.xlsx
python create_financieel_input.py

# Auto-update fuel/electricity prices from APIs
python update_financieel_input.py
```

### Fuel/Charging Station Fetcher

Fetches nearby fuel stations and EV chargers for each bus station:

```bash
# Auto-discover stations from input Excel
python fetch_tanklocaties.py --input Bijlage_J.xlsx

# With specific radius
python fetch_tanklocaties.py --input Bijlage_J.xlsx --radius 10

# Output: tanklocaties.json
```

## Future Versions (6-9)

See `PLAN_FINANCIAL_OPTIMIZATION.md` for planned financial optimization:

| Version | Description |
|---------|-------------|
| 6 | Financial overlay on existing roster (revenue, costs, profit calculation) |
| 7 | Euro-based cost optimization (driver costs, fuel costs, ORT surcharges) |
| 8 | Sustainability optimization (ZE/HVO100 fuel type assignment) |
| 9 | Full fuel logistics (fueling/charging schedule, station routing) |

## Command Line Reference

```
usage: busomloop_optimizer.py [-h] [--output OUTPUT] [--algoritme {greedy,mincost,all}]
                               [--deadhead DEADHEAD] [--snel]
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
  --snel                   Fast mode: skip non-greedy for versions 1-4
  --keer-dd MIN            Turnaround time Dubbeldekker (default: 8)
  --keer-tc MIN            Turnaround time Touringcar (default: 6)
  --keer-lvb MIN           Turnaround time Lagevloerbus (default: 5)
  --keer-midi MIN          Turnaround time Midi bus (default: 4)
  --keer-taxi MIN          Turnaround time Taxibus (default: 3)
```

## Example Workflow

```bash
# 1. Generate deadhead matrix (optional, enables version 5)
python google_maps_distances.py --input Bijlage_J.xlsx --key $GOOGLE_MAPS_API_KEY

# 2. Run optimizer with all versions
python busomloop_optimizer.py Bijlage_J.xlsx --deadhead deadhead_matrix.json --output project_x

# 3. Compare results
# - project_x_greedy_1_per_dienst.xlsx      → baseline
# - project_x_greedy_3_gecombineerd.xlsx    → cross-service optimization
# - project_x_greedy_5_deadhead.xlsx        → maximum optimization with repositioning

# 4. (Future) Prepare financial data for versions 7-9
python create_financieel_input.py
python update_financieel_input.py
python fetch_tanklocaties.py --input Bijlage_J.xlsx
```

## License

Internal use only — Gebo project for NS TVV optimization.
