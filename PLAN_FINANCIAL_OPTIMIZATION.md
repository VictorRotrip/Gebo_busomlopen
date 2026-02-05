# Plan: Financial Optimization for Bus Roster (Versions 6+)

## Context

The current optimizer (`busomloop_optimizer.py`) creates rosters in versions 1-5 for two
algorithms (greedy, mincost), optimizing primarily for **efficiency**: minimizing the number
of buses and idle time. This plan extends the optimizer with versions 6-9 that incorporate
**financial optimization**: maximizing profit by considering revenue, driver costs (CAO BB),
bus operating costs, sustainability bonuses, and penalty (malus) avoidance.

All financial variables are sourced from `Commerciele_variabelen.xlsx` (5 data sheets):
- **CAO_variabelen** — driver labor costs (80 variables)
- **Contract_variabelen** — contract terms, sustainability incentives (18 variables)
- **Malus_variabelen** — KPI penalties and bonuses (50 variables)
- **Prijzenblad_variabelen** — hourly rates, start fees (22 variables)
- **PvEisen_variabelen** — service level requirements (36 variables)

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
Parameters needed (not yet in spreadsheet — must be configured):
- Fuel consumption per bus type (L/100km for diesel/HVO, kWh/100km for ZE)
- Diesel B7 price per liter
- HVO100 price per liter
- Electricity price per kWh
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

- Modify min-cost flow: edge weights = `cost_saved − revenue_lost` instead of `deadhead × 2 + idle`
- Trade-offs: fewer expensive Dubbeldekkers if Touringcars suffice
- Minimize ORT exposure: prefer scheduling in daytime windows where possible
- Manage overtime: avoid shifts >7.5h to reduce break deductions and surcharges
- Allow bus type flexibility when capacity permits

### Version 8: Sustainability-Optimized (`8_duurzaamheid_geoptimaliseerd`)
**Goal**: Optimize ZE/HVO bus assignment to maximize sustainability bonuses and avoid malus

- Add fuel-type dimension to bus assignment (Diesel/HVO100/Electric)
- Calculate sustainability KPI scores and project malus/bonus
- Assign ZE buses to shorter routes (range-feasible) to maximize €0.12/km bonus
- Schedule HVO100 on remaining routes to hit sustainable fuel targets
- Determine optimal mix: the marginal cost of ZE vs. the bonus + malus avoidance

### Version 9: Full Financial + Fueling/Charging (`9_volledig_financieel`)
**Goal**: Full profit optimization including fuel logistics

- Range constraints for ZE buses; determine charging windows
- Fuel station locations for diesel/HVO along routes (Google Maps API)
- Charging station locations and availability
- Optimal fueling stops that minimize schedule disruption
- Complete profit optimization: revenue − all costs − penalties + all bonuses

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

### Step 2: Bus Operating Cost Config
New JSON config file (`bus_costs.json`) with:
```json
{
  "fuel_prices": {
    "diesel_b7_eur_per_liter": 1.85,
    "hvo100_eur_per_liter": 2.15,
    "electricity_eur_per_kwh": 0.25
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

### Step 3: Extend busomloop_optimizer.py
- New CLI flags: `--financieel`, `--bus-costs JSON`, `--brandstof-api`
- Import `financial_calculator` module
- Add version 6 output generation (financial overlay)
- Add version 7 optimization with profit objective
- Add version 8 with sustainability constraints
- Add version 9 with fueling/charging logistics

### Step 4: Extend Output Excel
New sheets per financial version:
- **Financieel Overzicht**: revenue, costs, profit per bus, per day, per bus type
- **CAO Kosten Detail**: paid hours, ORT, overtime, surcharges per shift
- **Duurzaamheid KPI**: sustainability scores, malus/bonus projection
- **Brandstof/Laden Plan**: fueling schedule, charging windows (version 9)

### Step 5: Optional API Integrations
- Fuel prices: [brandstof-check.nl API](https://brandstof-check.nl) or manual CSV
- Google Maps: extend `google_maps_distances.py` with fuel/charging station search
  - Places API: `type=gas_station` or `type=electric_vehicle_charging_station`
  - Filter along route corridors

---

## Variable Priority (Impact on Profit)

### Highest Impact
1. **Hourly rates** (Prijzenblad) — directly determines revenue
2. **Number of buses × active hours** — total revenue driver
3. **Bus type selection** — DD earns €36/h more than TC but costs more
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
2. **Dubbeldekker vs Touringcar**: higher revenue rate but higher operating cost
3. **Deadhead drives**: cost fuel/driver time but enable more revenue per bus
4. **Evening/night scheduling**: higher driver cost (ORT) but may be required
5. **Long shifts**: higher meal/break costs but fewer driver changeovers
6. **ZE buses**: lower fuel cost + €0.12/km bonus but range-limited
7. **HVO vs diesel**: higher fuel cost but up to €0.40/L incentive
8. **Sustainability target**: hitting targets avoids €1000/0.1pp malus

---

## Data Gaps to Resolve Before Implementation

| Gap | Action |
|-----|--------|
| Base hourly wage per driver | Get from CAO loonschalen or estimate |
| Fuel consumption per bus type | Get from fleet operator or estimate |
| Current fuel prices (diesel B7, HVO100, electricity) | Manual input or API |
| Bus fleet composition (how many ZE/HVO/diesel per type) | Get from operator |
| Charging infrastructure at stations | Survey or Google Maps API |
| Coordinator/traffic controller rates | Fill in Prijzenblad (currently €0) |
| Bus range and tank/battery capacity | Get from fleet specs |
