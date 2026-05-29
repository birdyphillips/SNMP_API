# Elements API - AQM Metrics Collection Training Guide

## Table of Contents

1. [Overview](#overview)
2. [What This Script Does](#what-this-script-does)
3. [Prerequisites](#prerequisites)
4. [Setup Instructions](#setup-instructions)
5. [How to Run the Script](#how-to-run-the-script)
6. [Understanding the Code](#understanding-the-code)
7. [Understanding the OIDs](#understanding-the-oids)
8. [Reading the Output](#reading-the-output)
9. [Troubleshooting](#troubleshooting)
10. [Key Concepts for Beginners](#key-concepts-for-beginners)

---

## Overview

This script uses the **Charter Elements API** to collect SNMP metrics from cable modems. It authenticates with the API, sends an SNMP walk request for a specific modem (by MAC address), and saves the results as a JSON file.

This is useful for:
- Collecting AQM (Active Queue Management) data from modems
- Gathering latency histogram bin data
- Monitoring QoS service flow statistics
- Validating LLD (Low Latency DOCSIS) configurations
- Troubleshooting modem performance issues

---

## What This Script Does

```
┌─────────────┐       ┌──────────────────┐       ┌──────────────┐
│  You run    │──────▶│  Elements API    │──────▶│  Cable Modem │
│  the script │       │  (authenticates) │       │  (SNMP walk) │
└─────────────┘       └──────────────────┘       └──────────────┘
                              │
                              ▼
                      ┌──────────────┐
                      │  JSON file   │
                      │  (results/)  │
                      └──────────────┘
```

**Step-by-step flow:**
1. Prompts you for a cable modem MAC address
2. Logs into the Elements API (gets an auth token)
3. Sends an SNMP walk request with a list of OIDs
4. Saves the response as a JSON file in `./results/`

---

## Prerequisites

| Requirement | Details |
|-------------|---------|
| Python | Version 3.8 or higher |
| Network | Must be on Charter/Spectrum network (VPN or on-prem) |
| Credentials | Your Charter AD username and password |
| Packages | `requests`, `python-dotenv` |

---

## Setup Instructions

### Step 1: Navigate to the Script Folder

```powershell
cd "C:\Users\<your_username>\OneDrive - Charter Communications\Access_Engineering_Project\Elements_API\elements_requests"
```

### Step 2: Install Python Dependencies

```powershell
pip install -r requirements.txt
```

This installs:
- `requests` — Makes HTTP calls to the API
- `python-dotenv` — Reads your credentials from the `.env` file

### Step 3: Configure Your Credentials

Edit the `.env` file in the same folder as the script:

```env
ELEMENTS_API_URL=https://api.elements.charter.com
ELEMENTS_API_USER=your_ad_username
ELEMENTS_API_PASS=your_ad_password
```

> ⚠️ **IMPORTANT:** Never commit the `.env` file to Git. It contains your password. The `.gitignore` file is already set up to exclude it.

### Step 4: Verify Setup

```powershell
python get_metrics_info.py
```

You should see:
```
Enter MAC address (any format):
```

---

## How to Run the Script

### Basic Usage

```powershell
python get_metrics_info.py
```

### Enter a MAC Address

The script accepts MAC addresses in **any format**:

| Format | Example |
|--------|---------|
| No separators | `08a7c0885eff` |
| Colon-separated | `08:a7:c0:88:5e:ff` |
| Dash-separated | `08-a7-c0-88-5e-ff` |
| Dot-separated (Cisco style) | `08a7.c088.5eff` |
| Uppercase | `08A7C0885EFF` |

All formats are normalized to lowercase with no separators before being sent to the API.

### Example Session

```
PS> python get_metrics_info.py
Enter MAC address (any format): 08:A7:C0:88:5E:FF
Login successful. Token expires: 2026-05-26T18:27:32.290529Z
Running SNMP on 08a7c0885eff...
  Done: 08a7c0885eff
Results are in the ./results directory
```

### Finding Your Results

Results are saved in the `./results/` folder with this naming format:
```
results/<mac_address>_<timestamp>.json
```

Example: `results/08a7c0885eff_2026-05-26T153000.json`

---

## Understanding the Code

### File Structure

```
elements_requests/
├── .env                  # Your credentials (DO NOT SHARE)
├── .gitignore            # Prevents .env from being committed
├── get_metrics_info.py   # The main script
├── requirements.txt      # Python package dependencies
├── results/              # Output folder (created automatically)
│   └── <mac>_<timestamp>.json
└── README.md             # This training document
```

### Code Breakdown

#### 1. Imports and Configuration

```python
from dotenv import load_dotenv  # Reads .env file
load_dotenv()                   # Loads .env variables into os.environ

BASE_URL = os.getenv('ELEMENTS_API_URL')  # Gets value from .env
USER = os.getenv('ELEMENTS_API_USER')
PASSWORD = os.getenv('ELEMENTS_API_PASS')
```

**What's happening:** The script reads your credentials from the `.env` file so you don't have to hardcode passwords in the script.

#### 2. Login Function

```python
def login():
    creds = b64encode(f"{USER}:{PASSWORD}".encode()).decode()
    resp = requests.post(
        f"{BASE_URL}/api/login/",
        headers={'Authorization': f'Basic {creds}'},
        json={},
        verify=False
    )
```

**What's happening:**
- Encodes `username:password` as base64 (this is how Basic Auth works)
- Sends a POST request to `/api/login/`
- `verify=False` skips SSL certificate checking (needed because Elements uses a self-signed cert)
- Returns a **token** that we use for all future requests

#### 3. SNMP Walk Function

```python
def snmp_walk(token, device_id):
    resp = requests.post(
        f"{BASE_URL}/api/snmp/walk/{device_id}?ensure_cm=true",
        headers={'Authorization': f'Token {token}'},
        json={"oids": AQM_OIDS},
        verify=False
    )
```

**What's happening:**
- Uses the token from login (not username/password)
- Sends the list of OIDs we want to query
- `ensure_cm=true` tells the API to verify this is a cable modem
- Returns all the SNMP data as a JSON dictionary

#### 4. Main Execution

```python
if __name__ == "__main__":
    raw = input("Enter MAC address (any format): ")
    device = normalize_mac(raw)
    token = login()
    gather_aqm_mibs_by_device(token, device)
```

**What's happening:**
- Prompts for MAC, normalizes it, logs in, runs the SNMP walk, saves results

---

## Understanding the OIDs

### What is an OID?

An **OID (Object Identifier)** is a unique address for a piece of data on a network device. Think of it like a file path on your computer, but for SNMP data on a modem.

Example: `1.3.6.1.4.1.4491.2.1.21.1.29.2.1.1` = Latency Max Latency value

### OID Categories in This Script

#### Device Information
| OID Name | What It Tells You |
|----------|-------------------|
| `firmware_current_version` | Current firmware running on the modem |
| `modem_ip` | IP address of the modem |
| `modem_mac` | MAC address of the modem |
| `modem_reg_state` | Registration state (is it online?) |
| `init_state` | Initialization state during boot |
| `docsis_base_capability` | DOCSIS version supported (3.0, 3.1, 4.0) |

#### Latency Histogram Bins
| OID Name | What It Tells You |
|----------|-------------------|
| `docsQosSfLatencyMaxLatency` | Maximum latency observed |
| `docsQosSfLatencyNumHistUpdates` | How many times the histogram was updated |
| `docsQosSfLatencyBin1Pkts` through `Bin16Pkts` | Packet counts in each latency bucket |

**Why this matters:** Latency bins show you the distribution of packet delays. Bin 1 = lowest latency, Bin 16 = highest. For LLD (Low Latency DOCSIS), you want most packets in the lower bins.

#### QoS Service Flow Info
| OID Name | What It Tells You |
|----------|-------------------|
| `qos_service_flow_direction` | 1 = Downstream, 2 = Upstream |
| `qos_service_flow_primary` | Is this the primary service flow? |
| `qos_service_flow_sid` | Service Flow ID number |
| `qos_service_flow_buffer_size` | Current buffer size in bytes |

#### QoS Parameter Set (Provisioned Settings)
| OID Name | What It Tells You |
|----------|-------------------|
| `qos_param_set_priority` | Traffic priority (0-7, higher = more priority) |
| `qos_param_set_max_traffic_rate` | Max speed in bits/sec (e.g., 1000000000 = 1 Gbps) |
| `qos_param_set_max_traffic_burst` | Max burst size in bytes |
| `qos_param_set_max_concat_burst` | Max concatenation burst (upstream) |
| `qos_param_set_service_class_name` | Service class name (e.g., "dsHSI029", "usHSI029") |
| `docsQosParamSetAqmLatencyTarget` | AQM latency target in microseconds |
| `docsQosParamSetMinimumBuffer` | Minimum buffer size |
| `docsQosParamSetTargetBuffer` | Target buffer size |
| `docsQosParamSetMaximumBuffer` | Maximum buffer size |

#### QoS Service Flow Statistics
| OID Name | What It Tells You |
|----------|-------------------|
| `qos_service_flow_pkts` | Total packets through this service flow |
| `qos_service_flow_octets` | Total bytes through this service flow |
| `qos_service_flow_policed_drop_pkts` | Packets dropped by traffic policing |
| `qos_service_flow_policed_delay_pkts` | Packets delayed by traffic policing |
| `qos_service_flow_aqm_dropped_pkts` | Packets dropped by AQM |

#### QoS Upstream Stats
| OID Name | What It Tells You |
|----------|-------------------|
| `qos_service_us_stats_tx_retries` | Transmission retry count |
| `qos_service_us_stats_tx_exceededs` | Transmission retry limit exceeded |
| `qos_service_us_stats_rq_retries` | Request retry count |
| `qos_service_us_stats_rq_exceededs` | Request retry limit exceeded |

#### Congestion Metrics (ECN/L4S)
| OID Name | What It Tells You |
|----------|-------------------|
| `docsQosSfCongestionAqmDroppedPkts` | Packets dropped by AQM congestion |
| `docsQosSfCongestionSanctionedPkts` | Packets sanctioned (rate limited) |
| `docsQosSfCongestionTotalEct0Pkts` | ECT(0) marked packets (Classic ECN) |
| `docsQosSfCongestionTotalEct1Pkts` | ECT(1) marked packets (L4S/LLD) |
| `docsQosSfCongestionCeMarkedEct1Pkts` | ECT(1) packets marked with CE |
| `docsQosSfCongestionArrivedCePkts` | Packets that arrived already CE-marked |

---

## Reading the Output

### Sample JSON Output Structure

```json
{
  "success": true,
  "data": [
    {
      "request_id": "08a7c0885eff",
      "ip": "10.x.x.x",
      "region": "stprmo",
      "response": {
        "Uptime": 1234567,
        "walked_data": {
          "1": {
            "docsQosSfLatencyMaxLatency": 5000,
            "docsQosSfLatencyBin1Pkts": 98234,
            "docsQosSfLatencyBin2Pkts": 1523,
            "qos_param_set_service_class_name": "dsHSI029",
            "qos_param_set_max_traffic_rate": 1000000000
          },
          "2": {
            "docsQosSfLatencyMaxLatency": 3200,
            "qos_param_set_service_class_name": "usHSI029",
            "qos_param_set_max_traffic_rate": 200000000
          }
        }
      }
    }
  ]
}
```

### How to Interpret

- **walked_data** is grouped by service flow index ("1", "2", etc.)
- Each index contains the OID values for that service flow
- Look at `qos_service_flow_direction`: 1 = DS (download), 2 = US (upload)
- `qos_param_set_max_traffic_rate` is in bits/sec (divide by 1,000,000 for Mbps)

---

## Troubleshooting

### Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| `ELEMENTS_API_USER and ELEMENTS_API_PASS must be set` | Missing `.env` file or empty values | Create/edit `.env` with your credentials |
| `SSL: CERTIFICATE_VERIFY_FAILED` | SSL verification enabled | Already handled with `verify=False` |
| `Login failed` | Wrong username/password | Check `.env` credentials |
| `Invalid MAC address` | MAC has wrong characters or length | Use a valid 12-character hex MAC |
| `Connection refused` | Not on Charter network | Connect to VPN |
| `ModuleNotFoundError: No module named 'dotenv'` | Missing dependency | Run `pip install python-dotenv` |
| `ModuleNotFoundError: No module named 'requests'` | Missing dependency | Run `pip install requests` |

### Verifying Your Environment

```powershell
# Check Python version (need 3.8+)
python --version

# Check if packages are installed
pip list | findstr requests
pip list | findstr dotenv

# Check if .env file exists
type .env
```

---

## Key Concepts for Beginners

### What is an API?

An **API (Application Programming Interface)** is a way for programs to talk to each other. Instead of opening a web browser and clicking buttons, our script sends HTTP requests directly to the Elements server and gets data back.

### What is REST?

The Elements API is a **REST API**, meaning:
- We use HTTP methods (GET, POST, PATCH) to interact with it
- Data is sent/received as JSON
- Each URL (endpoint) represents a resource

### What is Authentication?

1. **Basic Auth** — We send `username:password` encoded in base64. Used only for login.
2. **Token Auth** — After login, we get a token (like a temporary password). We use this for all other requests so we don't keep sending our real password.

### What is SNMP?

**SNMP (Simple Network Management Protocol)** is a standard protocol for monitoring network devices. Cable modems expose their stats via SNMP OIDs that we can query remotely through the Elements API.

### What is AQM?

**AQM (Active Queue Management)** is a technique used in DOCSIS networks to manage packet queues and reduce latency (bufferbloat). The metrics we collect help us verify AQM is working correctly on LLD-enabled modems.

### What is LLD?

**LLD (Low Latency DOCSIS)** is a feature in DOCSIS 3.1+ that reduces latency for real-time applications (gaming, video calls). It uses separate queues and ECN marking to achieve lower delays.

### What is ECN?

**ECN (Explicit Congestion Notification)** is a way for network devices to signal congestion without dropping packets. Instead of dropping, they mark packets with a CE (Congestion Experienced) flag so the sender can slow down.

- **ECT(0)** = Classic ECN traffic
- **ECT(1)** = L4S (Low Latency, Low Loss, Scalable throughput) traffic

---

## Quick Reference

### Run the Script
```powershell
python get_metrics_info.py
```

### Find Results
```powershell
dir results\
```

### View a Result File
```powershell
type results\08a7c0885eff_2026-05-26T153000.json
```

### Common MAC Address Sources
- CMTS: `show cable modem`
- Elements UI: Device lookup
- DLPQS: MAC search
- Modem label: Physical sticker on device

---

## Additional Resources

- Elements API Swagger Docs: `https://api.elements.charter.com/swagger/`
- DOCSIS 3.1 MIB Reference: CableLabs DOCS-QOS3-MIB
- LLD Chalk Page: See `Chalk_Pages/Low_Latency_DOCSIS_LLD.md`
- AQM Chalk Page: See `Chalk_Pages/Active_Queue_Management_AQM.md`
