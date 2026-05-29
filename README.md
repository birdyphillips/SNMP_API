# SNMP_API

Charter Access Engineering — API tools for collecting SNMP and QoS metrics from cable modems.

---

## Repository Structure

```
SNMP_API/
├── elements_api/          # Elements API — SNMP walk via Charter Elements API
│   ├── get_metrics_info.py
│   ├── requirements.txt
│   └── README.md
└── scope_api/             # Scope API — QoS service flow metrics & latency analysis
    ├── get_modem_metrics.py
    └── README.md
```

---

## elements_api

Connects to the Charter Elements API to perform SNMP walks on cable modems and collect AQM/LLD metrics (latency bins, service flow stats, congestion counters). Results are saved as JSON files.

See [elements_api/README.md](elements_api/README.md) for full setup and usage.

**Quick start:**
```bash
cd elements_api
pip install -r requirements.txt
# Add credentials to .env (see README)
python get_metrics_info.py
```

---

## scope_api

Connects to the Charter Scope API to collect QoS service flow metrics for one or more cable modems. Supports single snapshot, before/after delta, and continuous polling modes. Computes latency percentiles (P50/P99/P99.9) using linear interpolation and AVG methods.

See [scope_api/README.md](scope_api/README.md) for full usage and output examples.

**Quick start:**
```bash
cd scope_api
pip install requests python-dotenv
# Add credentials to .env (see README)
python get_modem_metrics.py --mac <modem_mac>
```

---

## Setup

Both tools require a `.env` file with API credentials. See each subfolder's README for the required variables. The `.env` file is excluded from version control via `.gitignore`.
