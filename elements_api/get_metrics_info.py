#!/usr/bin/env python3
"""
Elements API - AQM Metrics Collection Script
=============================================
This script connects to the Charter Elements API, authenticates,
and performs an SNMP walk on a cable modem to collect AQM (Active Queue Management)
metrics like latency bins, service flow stats, and congestion data.

Results are saved as JSON files in a ./results folder.

Prerequisites:
    1. Install dependencies:  pip install requests python-dotenv
    2. Create a .env file with your credentials (see .env in this folder)
"""

# ============================================================
# IMPORTS - Libraries we need to make this script work
# ============================================================
import os          # For reading environment variables and creating folders
import sys         # For exiting the script on errors
import json        # For converting Python data to/from JSON format
import requests    # For making HTTP requests to the API (like a web browser)
from base64 import b64encode  # For encoding username:password for Basic Auth
from datetime import datetime  # For timestamps on output files
from threading import Semaphore  # For limiting concurrent API calls
from concurrent.futures import ThreadPoolExecutor  # For running tasks in parallel
from dotenv import load_dotenv  # For reading credentials from .env file
import urllib3     # For suppressing SSL warnings

# ============================================================
# CONFIGURATION
# ============================================================

# Suppress SSL warnings (Elements API uses a self-signed certificate)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Load variables from .env file into environment
# This reads ELEMENTS_API_URL, ELEMENTS_API_USER, ELEMENTS_API_PASS
load_dotenv()

# Semaphore limits how many API calls can happen at the same time (max 25)
lock = Semaphore(25)

# Read credentials from .env file
BASE_URL = os.getenv('ELEMENTS_API_URL', 'https://<elements-api-host>')
USER = os.getenv('ELEMENTS_API_USER')
PASSWORD = os.getenv('ELEMENTS_API_PASS')

# Stop the script if credentials are missing
if not USER or not PASSWORD:
    print("ERROR: ELEMENTS_API_USER and ELEMENTS_API_PASS must be set in .env")
    sys.exit(1)


# ============================================================
# MAC ADDRESS HELPER
# ============================================================
def normalize_mac(mac):
    """
    Takes a MAC address in any format and returns it as lowercase
    with no separators (just 12 hex characters).

    Examples:
        "08:A7:C0:88:5E:FF" -> "08a7c0885eff"
        "08-a7-c0-88-5e-ff" -> "08a7c0885eff"
        "08a7.c088.5eff"    -> "08a7c0885eff"
        "08A7C0885EFF"      -> "08a7c0885eff"
    """
    return mac.strip().replace(":", "").replace("-", "").replace(".", "").lower()


# ============================================================
# HARDCODED OIDs (Object Identifiers)
# ============================================================
# These are the SNMP OIDs we want to query from the cable modem.
# Each OID represents a specific metric on the device.
# "oid_name" = human-readable name, "oid" = numeric SNMP identifier
AQM_OIDS = [
    # --- Latency Histogram Bins (how many packets fell into each latency bucket) ---
    {"oid_name": "docsQosSfLatencyMaxLatency",         "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.1"},
    {"oid_name": "docsQosSfLatencyNumHistUpdates",     "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.2"},
    {"oid_name": "docsQosSfLatencyBin1Pkts",           "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.3"},
    {"oid_name": "docsQosSfLatencyBin2Pkts",           "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.4"},
    {"oid_name": "docsQosSfLatencyBin3Pkts",           "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.5"},
    {"oid_name": "docsQosSfLatencyBin4Pkts",           "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.6"},
    {"oid_name": "docsQosSfLatencyBin5Pkts",           "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.7"},
    {"oid_name": "docsQosSfLatencyBin6Pkts",           "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.8"},
    {"oid_name": "docsQosSfLatencyBin7Pkts",           "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.9"},
    {"oid_name": "docsQosSfLatencyBin8Pkts",           "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.10"},
    {"oid_name": "docsQosSfLatencyBin9Pkts",           "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.11"},
    {"oid_name": "docsQosSfLatencyBin10Pkts",          "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.12"},
    {"oid_name": "docsQosSfLatencyBin11Pkts",          "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.13"},
    {"oid_name": "docsQosSfLatencyBin12Pkts",          "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.14"},
    {"oid_name": "docsQosSfLatencyBin13Pkts",          "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.15"},
    {"oid_name": "docsQosSfLatencyBin14Pkts",          "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.16"},
    {"oid_name": "docsQosSfLatencyBin15Pkts",          "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.17"},
    {"oid_name": "docsQosSfLatencyBin16Pkts",          "oid": "1.3.6.1.4.1.4491.2.1.21.1.29.2.1.18"},

    # --- Service Flow Buffer Size ---
    {"oid_name": "qos_service_flow_buffer_size",       "oid": "1.3.6.1.4.1.4491.2.1.21.1.3.1.17"},

    # --- Congestion Metrics (ECN/L4S related counters) ---
    {"oid_name": "docsQosSfCongestionAqmDroppedPkts",  "oid": "1.3.6.1.4.1.4491.2.1.21.1.30.1.1"},
    {"oid_name": "docsQosSfCongestionSanctionedPkts",  "oid": "1.3.6.1.4.1.4491.2.1.21.1.30.1.4"},
    {"oid_name": "docsQosSfCongestionTotalEct0Pkts",   "oid": "1.3.6.1.4.1.4491.2.1.21.1.30.1.5"},
    {"oid_name": "docsQosSfCongestionTotalEct1Pkts",   "oid": "1.3.6.1.4.1.4491.2.1.21.1.30.1.6"},
    {"oid_name": "docsQosSfCongestionCeMarkedEct1Pkts","oid": "1.3.6.1.4.1.4491.2.1.21.1.30.1.7"},
    {"oid_name": "docsQosSfCongestionArrivedCePkts",   "oid": "1.3.6.1.4.1.4491.2.1.21.1.30.1.8"},

    # --- Device Info (firmware, IP, MAC, registration state) ---
    {"oid_name": "firmware_current_version",           "oid": "1.3.6.1.2.1.69.1.3.5.0"},
    {"oid_name": "modem_ip",                           "oid": "1.3.6.1.4.1.4491.2.1.20.1.1.1.5"},
    {"oid_name": "modem_mac",                          "oid": "1.3.6.1.4.1.4491.2.1.20.1.1.1.2"},
    {"oid_name": "modem_reg_state",                    "oid": "1.3.6.1.4.1.4491.2.1.20.1.1.1.7"},
    {"oid_name": "init_state",                         "oid": "1.3.6.1.4.1.4491.2.1.20.1.1.1.9"},
    {"oid_name": "docsis_base_capability",             "oid": "1.3.6.1.4.1.4491.2.1.20.1.1.1.4"},

    # --- QoS Service Flow Info (direction, primary, SID) ---
    {"oid_name": "qos_service_flow_direction",         "oid": "1.3.6.1.4.1.4491.2.1.21.1.3.1.2"},
    {"oid_name": "qos_service_flow_primary",           "oid": "1.3.6.1.4.1.4491.2.1.21.1.3.1.3"},
    {"oid_name": "qos_service_flow_sid",               "oid": "1.3.6.1.4.1.4491.2.1.21.1.3.1.5"},

    # --- QoS Parameter Set (provisioned service flow settings) ---
    {"oid_name": "qos_param_set_priority",             "oid": "1.3.6.1.4.1.4491.2.1.21.1.2.1.5.2.1"},
    {"oid_name": "qos_param_set_max_traffic_rate",     "oid": "1.3.6.1.4.1.4491.2.1.21.1.2.1.6.2.1"},
    {"oid_name": "qos_param_set_max_traffic_burst",    "oid": "1.3.6.1.4.1.4491.2.1.21.1.2.1.7.2.1"},
    {"oid_name": "qos_param_set_max_concat_burst",     "oid": "1.3.6.1.4.1.4491.2.1.21.1.2.1.8.2.1"},
    {"oid_name": "qos_param_set_service_class_name",   "oid": "1.3.6.1.4.1.4491.2.1.21.1.2.1.4.2.2"},
    {"oid_name": "docsQosParamSetAqmLatencyTarget",    "oid": "1.3.6.1.4.1.4491.2.1.21.1.2.1.43.2.1"},
    {"oid_name": "docsQosParamSetMinimumBuffer",       "oid": "1.3.6.1.4.1.4491.2.1.21.1.2.1.39.2.1"},
    {"oid_name": "docsQosParamSetTargetBuffer",        "oid": "1.3.6.1.4.1.4491.2.1.21.1.2.1.40.2.1"},
    {"oid_name": "docsQosParamSetMaximumBuffer",       "oid": "1.3.6.1.4.1.4491.2.1.21.1.2.1.41.2.1"},

    # --- QoS Service Flow Stats (traffic counters) ---
    {"oid_name": "qos_service_flow_pkts",              "oid": "1.3.6.1.4.1.4491.2.1.21.1.4.1.1"},
    {"oid_name": "qos_service_flow_octets",            "oid": "1.3.6.1.4.1.4491.2.1.21.1.4.1.2"},
    {"oid_name": "qos_service_flow_policed_drop_pkts", "oid": "1.3.6.1.4.1.4491.2.1.21.1.4.1.6"},
    {"oid_name": "qos_service_flow_policed_delay_pkts","oid": "1.3.6.1.4.1.4491.2.1.21.1.4.1.7"},
    {"oid_name": "qos_service_flow_aqm_dropped_pkts", "oid": "1.3.6.1.4.1.4491.2.1.21.1.4.1.8"},

    # --- QoS Upstream Service Flow Stats (retry/exceeded counters) ---
    {"oid_name": "qos_service_us_stats_tx_retries",    "oid": "1.3.6.1.4.1.4491.2.1.21.1.5.1.1"},
    {"oid_name": "qos_service_us_stats_tx_exceededs",  "oid": "1.3.6.1.4.1.4491.2.1.21.1.5.1.2"},
    {"oid_name": "qos_service_us_stats_rq_retries",    "oid": "1.3.6.1.4.1.4491.2.1.21.1.5.1.3"},
    {"oid_name": "qos_service_us_stats_rq_exceededs",  "oid": "1.3.6.1.4.1.4491.2.1.21.1.5.1.4"},
]


# ============================================================
# API FUNCTIONS
# ============================================================

def login():
    """
    Authenticates with the Elements API using Basic Auth (username:password).
    Returns an auth token that we use for all subsequent API calls.

    How it works:
        1. Encode "username:password" in base64 format
        2. Send a POST request to /api/login/ with that encoded string
        3. The API returns a token we can use instead of sending credentials every time
    """
    # Encode credentials as base64 (required for Basic Authentication)
    creds = b64encode(f"{USER}:{PASSWORD}".encode()).decode()

    # Make the login request
    resp = requests.post(
        f"{BASE_URL}/api/login/",
        headers={
            'Content-Type': 'application/json;charset=UTF-8',
            'Authorization': f'Basic {creds}'
        },
        json={},          # Empty body - credentials are in the header
        verify=False      # Skip SSL verification (self-signed cert)
    )

    # Parse the JSON response
    data = resp.json()

    # Handle different response formats from the API
    if 'token' in data:
        print(f"Login successful. Token expires: {data['expiry']}")
        return data['token']
    elif data.get('success') and 'response' in data:
        print(f"Login successful. Token expires: {data['response']['expiry']}")
        return data['response']['token']
    else:
        print(f"Login failed: {data}")
        sys.exit(1)


def snmp_walk(token, device_id):
    """
    Performs an SNMP walk on a device via the Elements API.

    What is SNMP walk?
        SNMP (Simple Network Management Protocol) lets us query network devices
        for their stats. A "walk" means we're requesting multiple OIDs at once
        and getting back all matching values.

    Args:
        token:     The auth token from login()
        device_id: The MAC address of the cable modem (no separators, lowercase)

    Returns:
        A dictionary containing all the SNMP data from the device
    """
    resp = requests.post(
        f"{BASE_URL}/api/snmp/walk/{device_id}?ensure_cm=true",
        headers={
            'Content-Type': 'application/json;charset=UTF-8',
            'Authorization': f'Token {token}'  # Use token auth (not Basic)
        },
        json={"oids": AQM_OIDS},  # Send our list of OIDs to query
        verify=False               # Skip SSL verification
    )
    return resp.json()


def gather_aqm_mibs_by_device(token, device):
    """
    Queries a single device and saves the results to a JSON file.

    Args:
        token:  The auth token from login()
        device: The MAC address to query
    """
    device = device.strip()

    # Semaphore ensures we don't exceed 25 simultaneous API calls
    with lock:
        # Call the SNMP walk API
        oid_info = snmp_walk(token, device)

        # Create a timestamp for the filename (e.g., 2026-05-26T153000)
        timestamp = datetime.now().strftime("%Y-%m-%dT%H%M%S")

        # Save results to a JSON file in the results/ folder
        try:
            with open(f'results/{device}_{timestamp}.json', 'w') as out_file:
                out_file.write(json.dumps(oid_info, indent=2))
            print(f"  Done: {device}")
        except Exception as err:
            print(f"Error writing results for {device}: {err}")


# ============================================================
# MAIN - Script starts here when you run it
# ============================================================
if __name__ == "__main__":
    # Prompt the user for a MAC address
    raw = input("Enter MAC address (any format): ")

    # Clean up the MAC address to a standard format
    device = normalize_mac(raw)

    # Validate: must be exactly 12 hex characters after normalization
    if len(device) != 12 or not all(c in '0123456789abcdef' for c in device):
        print(f"Invalid MAC address: {raw}")
        sys.exit(1)

    # Step 1: Login to get our auth token
    token = login()

    # Step 2: Create results folder if it doesn't exist
    try:
        os.mkdir('results')
    except FileExistsError:
        pass  # Folder already exists, that's fine

    # Step 3: Run the SNMP walk and save results
    print(f"Running SNMP on {device}...")
    gather_aqm_mibs_by_device(token, device)
    print("Results are in the ./results directory")
