# Plan: Financial Optimization for Bus Roster (Versions 6+)

## Context

The current optimizer (`busomloop_optimizer.py`) creates rosters in versions 1-5 for two
algorithms (greedy, mincost), optimizing primarily for **efficiency**: minimizing the number
of buses and idle time. This plan extends the optimizer with versions 6-9 that incorporate
**financial optimization**: maximizing profit by considering revenue, driver costs (CAO BB),
bus operating costs, sustainability bonuses, and penalty (malus) avoidance.

**Important constraint**: Bus types are dictated by NS in the input Excel (Bijlage J). If NS
specifies a Dubbeldekker for a trip, we must use a Dubbeldekker. The optimizer cannot swap bus
types. What the optimizer CAN optimize is:
- Which trips to chain together on the same bus (affects idle time, deadhead, shift length)
- Which fuel type to use per bus (diesel B7 / HVO100 / electric ZE)
- How to schedule fueling/charging stops
- How to minimize driver costs through smart chaining (reducing ORT, overtime, broken shifts)

All financial variables are sourced from `Commerciele_variabelen.xlsx` (5 data sheets):
- **CAO_variabelen** — driver labor costs (80 variables)
- **Contract_variabelen** — contract terms, sustainability incentives (18 variables)
- **Malus_variabelen** — KPI penalties and bonuses (50 variables)
- **Prijzenblad_variabelen** — hourly rates, start fees (22 variables)
- **PvEisen_variabelen** — service level requirements (36 variables)

---

## How the Current Mincost Algorithm Works

### Network Structure
The algorithm models trip chaining as a **bipartite matching problem**:
- **Nodes**: Each of the `n` trips is both a potential predecessor (left) and successor (right)
- **Edges**: `i → j` exists if trip `i` can be followed by trip `j` on the same bus
- **Edge cost**: `deadhead_time × 2 + idle_time` (or just `idle_time` if same location)

### can_connect() Checks
Before an edge is created, ALL of these must hold:
1. Same bus type (`trip_i.bus_type == trip_j.bus_type`)
2. Same date (`trip_i.date_str == trip_j.date_str`)
3. Time order (`trip_j.departure > trip_i.arrival`)
4. Location feasible (same station, or deadhead matrix allows travel)
5. Time feasible (gap ≥ turnaround time + deadhead driving time)

### Cost Function (Current)
```
idle = trip_j.departure - trip_i.arrival
cost = deadhead_time × 2 + idle    (if different locations)
cost = idle                         (if same location)
```
Deadhead is weighted 2× because empty driving wastes fuel + driver time with no revenue.

### SPFA Algorithm
The Successive Shortest Path Algorithm (SPFA) finds:
1. **Maximum matching** = minimum number of buses (primary objective)
2. **Minimum total cost** = least deadhead + idle (secondary objective)

It works by repeatedly finding the cheapest "augmenting path" — a chain of reassignments
that merges two buses into one — until no more merges are possible.

### Result
Chains like `[A→B→D]` and `[C→E]`, where each chain = one BusRotation (one bus's daily
schedule). The algorithm guarantees the minimum possible number of buses.

### What Changes for Financial Optimization
The cost function changes from `deadhead × 2 + idle` to a euro-based cost that considers:
- Revenue lost during idle/deadhead time (bus earns nothing)
- Driver cost during idle time (still getting paid)
- Fuel cost during deadhead drives
- ORT surcharges if chaining extends shifts into unsocial hours
- Break deduction impacts from shift length changes

---

## Core Financial Formula

```
PROFIT = REVENUE − DRIVER_COSTS − BUS_OPERATING_COSTS − PENALTIES + BONUSES
```

### Revenue
```
REVENUE = Σ (active_driving_hours_per_bus × hourly_rate_per_bustype)
```
Only earned during active driving + passenger loading/unloading. NOT during idle or deadhead.

| Bus Type         | Rate (€/h)  |
|------------------|-------------|
| Dubbeldekker     | €116.37     |
| Touringcar       | €80.455     |
| Lagevloer/gelede | €80.445     |
| Midi bus         | €74.85      |
| Taxibus          | €50.455     |

### Driver Costs (CAO BB)
```
DRIVER_COST_PER_SHIFT =
  (paid_hours × base_hourly_wage)
  + ORT_surcharge (€4.80–6.68/h for unsocial hours)
  + overtime_surcharge (35% riding, 30–100% non-riding)
  + break_surcharge (€14.72–15.92 per broken shift)
  + meal_allowance (€22.50–36.72 for shifts ≥11h/≥14h)
  × (1 + vakantietoeslag 8%)
```

Key CAO rules affecting paid hours:
- Break deduction brackets (Pauzestaffel):
  - ≤4.5h → 0 min, ≤7.5h → 30 min, ≤10.5h → 60 min
  - ≤13.5h → 90 min, ≤16.5h → 120 min, ≥16.5h → 150 min
- 1× unpaid break up to 1h per shift (OV rule)
- Night work (01:00–05:00): max 12h in 24h period
- Overtime threshold: 173.33h/month or 160h/4-weeks

### Bus Operating Costs
```
BUS_COST = Σ (km_driven × fuel_consumption_per_km × fuel_price)
         + deadhead_km × fuel_consumption × fuel_price
```
Parameters needed (from API or manual config):
- Fuel consumption per bus type (L/100km for diesel/HVO, kWh/100km for ZE)
- Diesel B7 price per liter (from CBS OData API)
- HVO100 price per liter (manual config from Rolande)
- Electricity price per kWh (from EnergyZero API)
- Bus range (km) for ZE, tank size for diesel
- Charging time for ZE buses

### Penalties (Malus)
```
MALUS = sustainability_fuel_malus + emission_malus + delivery_malus
```
- Sustainable fuel: €1,000 per 0.1pp below target (35% yr1 → 75% yr8)
- Emission norm: €1,000 per 0.1pp below 100% Euro VI/ZE
- Planned delivery: 0.1% annual revenue per pp below 95%
- Unplanned delivery: €5,000 per pp below 90%
- Total non-delivery: €30,000 flat
- **Total malus cap: 1% of annual revenue**

### Bonuses
```
BONUS = ZE_km × €0.12 + HVO_liters × min(price_diff, €0.35) + HVO_liters × €0.05
```
- Zero emission: €0.12/km
- HVO100 incentive: up to €0.40/liter (price diff capped at €0.35 + €0.05 stimulans)
- KPI bonuses for exceeding targets (max 0.5% annual revenue)

---

## Proposed New Versions

### Version 6: Financial Cost Calculation (`6_financieel_basis`)
**Goal**: Calculate full financial picture on existing efficient roster (from version 5)

- Input: version 5 roster + Commerciele_variabelen.xlsx + fuel config
- Per bus rotation: revenue (active hours × rate), driver cost (CAO), fuel cost
- Output: new Excel sheets "Financieel Overzicht", "CAO Kosten", "Duurzaamheid"
- No re-optimization; purely analytical overlay on the efficient roster

### Version 7: Cost-Optimized Roster (`7_financieel_geoptimaliseerd`)
**Goal**: Re-optimize with profit as objective instead of bus count

- Modify min-cost flow: edge weights become euro-based costs instead of `deadhead × 2 + idle`
- New cost per edge (i→j): `driver_cost(shift_with_j) - driver_cost(shift_without_j) + fuel_cost(deadhead) - revenue_from_reduced_buses`
- Minimize ORT exposure: prefer chaining trips that keep shifts within daytime hours
- Manage overtime: consider break deduction bracket jumps when extending shifts
- **Bus types remain fixed** as dictated by NS input — only chaining order changes

### Version 8: Sustainability-Optimized (`8_duurzaamheid_geoptimaliseerd`)
**Goal**: Optimize fuel-type assignment per bus to maximize sustainability bonuses and avoid malus

- Add fuel-type dimension to each bus rotation (Diesel B7 / HVO100 / Electric ZE)
- Bus type stays as NS specifies; fuel type is our choice
- Calculate sustainability KPI scores and project malus/bonus
- Assign ZE to shorter rotations (range-feasible) to maximize €0.12/km bonus
- Assign HVO100 to remaining rotations to hit sustainable fuel KPI targets
- Determine optimal mix: marginal cost of ZE/HVO vs. bonus + malus avoidance

### Version 9: Full Financial + Fueling/Charging (`9_volledig_financieel`)
**Goal**: Full profit optimization including fuel logistics

- Range constraints for ZE buses; determine charging windows during idle periods
- Fuel station locations for diesel/HVO along routes (OpenStreetMap + Google Maps)
- Charging station locations and availability (Open Charge Map API)
- Optimal fueling stops that minimize schedule disruption
- Complete profit optimization: revenue − all costs − penalties + all bonuses

---

## API Integrations for Automated Data

### Recommended API Stack (all free or freemium)

| Data Need | API | Free? | Auth | Notes |
|-----------|-----|-------|------|-------|
| Diesel B7 price | **CBS OData** (table `80416ned`) | Yes | None | Daily NL average, official government data |
| Electricity price | **EnergyZero API** | Yes | None | Hourly Dutch EPEX spot prices incl. BTW |
| HVO100 price | **Manual config** (Rolande website) | N/A | N/A | No public API exists for HVO100 |
| Fuel station locations | **OpenStreetMap Overpass** | Yes | None | Full NL, includes `fuel:HVO100=yes` tags |
| Fuel station prices | **HERE Fuel Prices API** | 250K/mo free | API key | Per-station real-time prices |
| EV charging stations | **Open Charge Map** | Yes | Free API key | Connector types, power ratings, locations |
| EV charging (premium) | **Eco-Movement** | Paid | Enterprise | 99% NL coverage, real-time availability |
| EU-level fuel reference | **EU Weekly Oil Bulletin** | Yes | None | Weekly national averages for cross-check |

### CBS OData — Diesel B7 Prices
```python
# pip install cbsodata pandas
import cbsodata, pandas as pd
data = pd.DataFrame(cbsodata.get_data('80416ned'))  # Daily pump prices NL
# Returns: Euro95, Diesel B7, LPG — weighted average incl. VAT
```
- URL: https://opendata.cbs.nl/ODataFeed/odata/80416ned
- License: CC BY 4.0
- No rate limits, no auth

### EnergyZero — Hourly Electricity Prices
```python
# pip install energyzero
from energyzero import EnergyZero
from datetime import date
import asyncio

async def get_prices():
    async with EnergyZero() as client:
        prices = await client.energy_prices(start_date=date(2026, 2, 4), end_date=date(2026, 2, 5))
        print(f"Current: {prices.current_price}, Avg: {prices.average_price}")
asyncio.run(get_prices())
```
- URL: https://api.energyzero.nl/v1/energyprices
- Dutch EPEX spot day-ahead, including BTW
- No auth required

### OpenStreetMap Overpass — Fuel & Charging Stations
```
# All fuel stations in NL (including HVO100 tags)
[out:json][timeout:25];
area["name"="Nederland"]->.searchArea;
node["amenity"="fuel"](area.searchArea);
out body;

# All EV charging stations in NL
[out:json][timeout:25];
area["name"="Nederland"]->.searchArea;
node["amenity"="charging_station"](area.searchArea);
out body;
```
- URL: https://overpass-api.de/api/interpreter
- Returns: lat/lon, brand, fuel types (fuel:diesel, fuel:HVO100), opening hours
- Completely free, no auth

### Open Charge Map — EV Charging Details
```
GET https://api.openchargemap.io/v3/poi/?output=json
    &latitude=52.3676&longitude=4.9041
    &distance=10&maxresults=50
    &countrycode=NL
    &key={FREE_API_KEY}
```
- Register at openchargemap.org for free API key
- Returns: location, operator, connector types (Type 2, CCS, CHAdeMO), power (kW), status

### HVO100 & B7 Diesel Advisory Prices — Contract-Specified Sources

The NS contract specifies that the monthly HVO100/B7 price difference is calculated from
the **landelijke adviesprijzen** (national advisory prices) of three commercial suppliers,
taken on the **first day of each month**. The price difference is the average across the
three sources.

**Contract formula:**
```
monthly_price_diff = average(
    (HVO100_price - B7_price) for PK_Energy,
    (HVO100_price - B7_price) for Fieten_Olie,
    (HVO100_price - B7_price) for BP_Nederland
)
# All prices taken on the 1st of the month

if price_diff > 0:
    incentive_per_liter = min(price_diff, €0.35) + €0.05   # capped at €0.40 total
else:
    incentive_per_liter = €0.00   # everything falls away (both diff and stimulans)
```

**Note**: If a source publishes per 100 liters, convert to per-liter first.

| Source | HVO100 | B7 Diesel | Format | URL | Automatable? |
|--------|--------|-----------|--------|-----|-------------|
| **Fieten Olie** | Yes | Yes | Per 100L, excl.+incl. BTW, daily, 1yr history | https://www.fieten.info/adviesprijzen/ | Best: daily updates, both fuels, historical data. Needs web scraper. |
| **PK Energy** | Yes | Yes | Per 100L, excl. BTW, PDF history | https://pkenergy.nl/brandstofprijzen/ | Medium: website + PDF. No API. |
| **BP Nederland** | Unclear | Yes | Per liter, incl. BTW | bp.com landelijke-adviesprijzen | Difficult: HVO100 not on standard advisory page. May need direct contact. |

**Current example prices (Feb 2026, Fieten Olie):**
- HVO100: €180.91 / 100L excl. BTW = €1.81/L
- Diesel B7: €167.69 / 100L excl. BTW = €1.68/L
- Price difference: €0.13/L → incentive = €0.13 + €0.05 = **€0.18/L**

**Additional HVO100 sources (for cross-reference):**
| Source | URL | Data |
|--------|-----|------|
| Rolande | https://rolande.eu/en/pricing/ | Published HVO100 prices excl. VAT |
| glpautogas.info | https://glpautogas.info/en/hvo100-stations-netherlands.html | 131 HVO100 stations in NL with prices |
| evofenedex | evofenedex.nl (members only since Nov 2024) | Weekly HVO diesel price overview |

### Alternative Fuel Station APIs (if more data needed)
- **ANWB API** (unofficial): ~3800 stations with prices, reverse-engineered from app
  - GitHub: https://github.com/bartmachielsen/ANWB-Fuel-Prices
- **HERE Fuel Prices**: per-station real-time prices, 250K free requests/month
- **Google Maps Places API**: use `type=gas_station` or `type=electric_vehicle_charging_station`
  - Already integrated via `google_maps_distances.py` — can extend with Places Nearby Search

---

## Implementation Steps

### Step 1: Financial Calculator Module (`financial_calculator.py`)
New module that reads Commerciele_variabelen.xlsx and exposes:
- `load_financial_config(xlsx_path) → dict`
- `calculate_driver_cost(rotation, date, config) → CostBreakdown`
- `calculate_revenue(rotation, bus_type, config) → float`
- `calculate_sustainability_score(rotations, config) → KPIScores`
- `calculate_malus_bonus(kpi_scores, annual_revenue, config) → MalusBonus`
- `calculate_profit(rotations, config) → ProfitReport`

Key logic:
- CAO break deduction brackets
- ORT time-window overlap calculation (minutes in each ORT window per shift)
- Overtime detection and surcharge calculation
- Date-dependent rate selection (2025-01-01, 2025-07-01, 2026-01-01)

### Step 2: Fuel Price Fetcher Module (`fuel_prices.py`)
New module that fetches fuel prices and station data:
- `fetch_diesel_b7_price() → float` — CBS OData API (daily NL average)
- `fetch_electricity_price(date, hour) → float` — EnergyZero API (hourly EPEX spot)
- `fetch_hvo100_incentive(month, year) → HVOIncentive` — scrape Fieten Olie + PK Energy + BP
  - Gets HVO100 and B7 prices from all 3 contract-specified sources on 1st of month
  - Calculates average price difference across sources
  - Applies contract formula: incentive = min(diff, €0.35) + €0.05 if diff > 0, else €0
- `fetch_fuel_stations(lat, lon, radius) → list` — OpenStreetMap Overpass
- `fetch_charging_stations(lat, lon, radius) → list` — Open Charge Map
- Cache results to avoid excessive API calls (fuel prices: 24h cache, stations: 7d cache)

### Step 3: Bus Operating Cost Config
New JSON config file (`bus_costs.json`) with:
```json
{
  "fuel_prices": {
    "diesel_b7_eur_per_liter": "auto",
    "hvo100_eur_per_liter": 2.15,
    "electricity_eur_per_kwh": "auto"
  },
  "consumption_per_100km": {
    "Dubbeldekker": {"diesel": 45, "hvo100": 45, "electric_kwh": 180},
    "Touringcar": {"diesel": 32, "hvo100": 32, "electric_kwh": 130},
    "Lagevloerbus": {"diesel": 38, "hvo100": 38, "electric_kwh": 150},
    "Midi bus": {"diesel": 25, "hvo100": 25, "electric_kwh": 100},
    "Taxibus": {"diesel": 12, "hvo100": 12, "electric_kwh": 50}
  },
  "range_km": {
    "Dubbeldekker_ZE": 250,
    "Touringcar_ZE": 300,
    "Lagevloerbus_ZE": 280,
    "Midi bus_ZE": 350,
    "Taxibus_ZE": 400
  },
  "base_hourly_wage_eur": 18.50,
  "driver_overhead_factor": 1.35
}
```
Values marked `"auto"` are fetched from APIs; others are manually configured.

### Step 4: Extend busomloop_optimizer.py
- New CLI flags: `--financieel`, `--bus-costs JSON`, `--brandstof-api`
- Import `financial_calculator` and `fuel_prices` modules
- Add version 6 output generation (financial overlay)
- Add version 7 optimization with euro-based cost function
- Add version 8 with sustainability fuel-type assignment
- Add version 9 with fueling/charging logistics

### Step 5: Extend Output Excel
New sheets per financial version:
- **Financieel Overzicht**: revenue, costs, profit per bus, per day, per bus type
- **CAO Kosten Detail**: paid hours, ORT, overtime, surcharges per shift
- **Duurzaamheid KPI**: sustainability scores, malus/bonus projection
- **Brandstof/Laden Plan**: fueling schedule, charging windows (version 9)

---

## Variable Priority (Impact on Profit)

### Highest Impact
1. **Hourly rates** (Prijzenblad) — directly determines revenue; fixed by NS per bus type
2. **Number of buses × active hours** — total revenue driver
3. **Shift composition** — which trips chain together affects idle time (= unpaid but costly)
4. **ORT surcharges** — €4–7/h extra cost during unsocial hours
5. **Sustainability KPI** — malus €1,000 per 0.1pp, adds up fast

### Medium Impact
6. **Overtime management** — 35% surcharge above threshold
7. **Break deductions** — determines paid vs unpaid hours
8. **Fuel costs** — depends on km driven and fuel type
9. **ZE bonus** (€0.12/km) and **HVO incentive** (up to €0.40/L)

### Lower Impact (contractually required)
10. Meal allowances — €22–37 for long shifts
11. Break surcharges — €10–16 per broken shift
12. Rest day compensation — 135% penalty
13. Coordinator/traffic controller costs — currently €0

---

## Key Trade-offs the Optimizer Must Balance

1. **More buses = more cost, but also more revenue** (if active hours are filled)
2. **Chaining order affects shift length** — longer shifts cross break deduction brackets and may trigger overtime
3. **Deadhead drives**: cost fuel/driver time but enable more revenue per bus (fewer buses needed)
4. **Evening/night scheduling**: higher driver cost (ORT) but may be required by NS timetable
5. **Long shifts**: higher meal/break costs but fewer driver changeovers needed
6. **ZE buses**: lower fuel cost + €0.12/km bonus but range-limited and need charging windows
7. **HVO vs diesel**: higher fuel cost but up to €0.40/L incentive
8. **Sustainability target**: hitting targets avoids €1000/0.1pp malus — worth investing in ZE/HVO

---

## Data Gaps to Resolve Before Implementation

| Gap | Action | Source | Status |
|-----|--------|--------|--------|
| Base hourly wage per driver | Get from CAO loonschalen or estimate | CAO Besluit Personenvervoer | Manual input needed |
| Fuel consumption per bus type | Get from fleet operator or estimate | Fleet data / manufacturer specs | Manual input needed |
| Current diesel B7 price | **Automated**: CBS OData API | https://opendata.cbs.nl (table 80416ned) | Ready to implement |
| Current electricity price | **Automated**: EnergyZero API | https://api.energyzero.nl | Ready to implement |
| HVO100/B7 price diff (incentive) | **Semi-automated**: scrape contract sources | Fieten Olie, PK Energy, BP Nederland (1st of month) | Sources identified |
| HVO100 price (absolute) | Scrape Fieten Olie adviesprijzen | https://www.fieten.info/adviesprijzen/ | Source found, needs scraper |
| Bus fleet composition (ZE/HVO/diesel) | Get from operator | Fleet inventory | Manual input needed |
| Fuel station locations | **Automated**: OpenStreetMap Overpass | https://overpass-api.de | Ready to implement |
| Charging station locations | **Automated**: Open Charge Map | https://openchargemap.org | Ready to implement |
| Coordinator/traffic controller rates | Fill in Prijzenblad (currently €0) | Contract negotiation | Manual input needed |
| Bus range and tank/battery capacity | Get from fleet specs | Manufacturer data | Manual input needed |
