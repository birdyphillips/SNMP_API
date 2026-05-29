# Scope API — Modem Metrics & Latency Report Tool

Collects QoS service flow metrics from the Scope API for one or more cable modems,
computes latency bin statistics, and prints a formatted summary to the console.

---

## Requirements

```
pip install requests
```

---

## Usage

```
python get_modem_metrics.py --mac <MAC> [--mac <MAC> ...] [--wait <seconds>] [--interval <seconds>]
```

### Arguments

| Argument | Description | Default |
|---|---|---|
| `--mac` | Modem MAC address — repeat for multiple modems | `08a7c0885eff` |
| `--wait` | Collect a before snapshot, wait N seconds, collect after, print delta | `0` (disabled) |
| `--interval` | Continuously poll every N seconds, printing delta each cycle — Ctrl+C to stop | `0` (disabled) |

### Examples

Single snapshot:
```
python get_modem_metrics.py --mac 08a7c0885eff
```

Before/after with 60 second wait:
```
python get_modem_metrics.py --mac 08a7c0885eff --wait 60
```

Before/after with 2 minute wait:
```
python get_modem_metrics.py --mac 08a7c0885eff --wait 120
```

Continuous polling every 60 seconds (Ctrl+C to stop):
```
python get_modem_metrics.py --mac 08a7c0885eff --interval 60
```

Continuous polling every 2 minutes, multiple modems:
```
python get_modem_metrics.py --mac 08a7c0885eff --mac a0ed6dff2890 --interval 120
```

---

## Modes

### Single Snapshot (`no --wait or --interval`)
Collects one snapshot and prints raw JSON + formatted summary with raw counters.

### Before/After (`--wait N`)
1. Collects **before** snapshot — prints raw JSON
2. Counts down N seconds
3. Collects **after** snapshot — prints raw JSON
4. Computes deltas and prints delta summary

### Continuous Polling (`--interval N`)
- Cycle 1 — collects first snapshot, prints summary with raw counters
- Cycle 2+ — collects new snapshot, computes delta against previous cycle, prints delta summary
- Counts down between each poll
- Press **Ctrl+C** to stop

---

## Console Output

```
================================================================================
  MODEM DELTA SUMMARY  —  08A7C0885EFF  (wait=60s)
================================================================================
  IP            : 192.168.1.1         DOCSIS  : 3.1
  Reg State     : operational         Init    : online
  Firmware      : ET2251-UNI-01.09.10-C3R
  Collected     : 2026-05-22 10:05:42

  ────────────────────────────────────────────────────────────────────────────
  SID 93890  [US]  usCBD011  (primary)  MaxRate=42 Mbps
  ────────────────────────────────────────────────────────────────────────────
  FLOW STATS DELTA
    Packets            :         16,842,507    Octets             :   10,057,348,219
    AQM Drops          :              1,532    Policed Drops      :               43
    ECT1 Pkts          :                  0    CE Marked          :                0

  LATENCY DELTA  (16,841,846 bin pkts)
    Metric                  Interpolation        AVG Method
    ----------------------------------------------------------
    Weighted Avg              1.6452 ms
    P50                       0.3049 ms           0.3750 ms
    P99                      24.4906 ms          25.0000 ms
    P99.9                    85.4598 ms          75.0000 ms
    ----------------------------------------------------------
    Max Latency             244,868 ms    Hist Updates  : 223,113,322

  SID 93891  [DS]  dsCBD011  (primary)  MaxRate=900 Mbps
  ────────────────────────────────────────────────────────────────────────────
    No traffic delta
================================================================================
```

---

## Latency Bin Edges

Upstream (US):
```
[0, 0.05, 0.10, 0.25, 0.50, 1.00, 2.00, 5.00, 10.00, 20.00, 30.00, 40.00, 50.00, 100.00, 150.00, 200.00, 500.00]
```

Downstream (DS):
```
[0, 0.05, 0.10, 0.25, 0.50, 1.00, 2.00, 5.00, 10.00, 20.00, 30.00, 40.00, 50.00, 60.00, 70.00, 80.00, 500.00]
```

Direction is auto-detected from `qos_service_flow_direction` (`"upstream"` / `"downstream"`).

---

## Latency Calculations

| Metric | Method |
|---|---|
| Weighted Avg | `sum(delta_i × bin_midpoint_i) / total` |
| P50 / P99 / P99.9 (Interpolation) | Linear interpolation within the bin where cumulative count crosses the target |
| P50 / P99 / P99.9 (AVG Method) | Bin midpoint of the first bin where cumulative count >= target |

---

## API Response Structure

The Scope API returns each metric as a nested dict keyed by `"2.<sfid>"`:

```json
{
  "qos_service_flow_direction": {
    "2.93890": "upstream",
    "2.93891": "downstream"
  },
  "docsQosSfLatencyBin1Pkts": {
    "2.93890": 57968244,
    "2.93892": 2
  }
}
```

Service class names use `"2.1.<sfid>"` keys.
Max traffic rate uses `"2.<sfid>.1"` keys.

---

## Fields Collected

- Firmware version, modem IP, reg state, DOCSIS capability
- Per-SF: direction, primary/secondary, service class, max rate
- Per-SF: packets, octets, AQM drops, policed drops, buffer size
- Per-SF: latency bins 1–16, max latency, histogram updates
- Per-SF: congestion ECT0/ECT1/CE-marked/sanctioned packets
- Per-SF: US stats — TX retries/exceededs, RQ retries/exceededs

---

## Project Structure

```
Scope_API/
├── get_modem_metrics.py   # Main script — API collection + console summary
├── snmp_parser.py         # Standalone pure-Python SNMP text file parser (no imports)
└── README.md              # This file
```

### Related Tools

| Path | Description |
|---|---|
| `Scripts/snmp_collector.py` | SSH jump server SNMP collection for iCMTS/vCMTS |
| `modem_metrics/get_snmp.py` | SNMP-based modem metrics collection |
| `Scripts/latency_calc.py` | Standalone latency bin calculator |
| `templates/CMTS_LATENCY_CALCULATOR.xlsx` | Manual latency bin Excel template |
