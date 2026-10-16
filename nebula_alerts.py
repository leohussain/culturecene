import multiprocessing
import json
import os
import tempfile
import atexit
import requests
from datetime import datetime, date

import urllib3
import warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
warnings.filterwarnings("ignore", message="Unverified HTTPS request")

# =============================================================================
# PASSWORD HANDLING
# =============================================================================

_env_password = os.environ.get("ARCGIS_PASSWORD")
if _env_password:
    PASSWORD = 'n0h47Rd6AWA0obOCNgvI!'
else:
    import getpass
    PASSWORD = getpass.getpass("Enter ArcGIS password (once only): ")

_PASS_FILE = os.path.join(tempfile.gettempdir(), "nebula_arcgis.tmp")

with open(_PASS_FILE, "w") as _f:
    _f.write(PASSWORD)

def _cleanup_pass_file():
    try:
        if os.path.exists(_PASS_FILE):
            os.remove(_PASS_FILE)
    except Exception:
        pass

atexit.register(_cleanup_pass_file)

# =============================================================================
# CONFIGURATION
# =============================================================================

PORTAL_URL = "https://axaxl-gis.cloud.esriuk.com/portal"
USERNAME   = "builtin.admin"

PORTFOLIO_LAYERS = {
    "Florida WS": {
        "url":        "https://axaxl-gis.cloud.esriuk.com/server/rest/services/Hosted/Florida_Aug26_Peril2/FeatureServer/0",
        "name_field": "AccountName",
        "id_field":   "Locid"
    },
    # "Australia WF": {
    #     "url":        "https://axaxl-gis.cloud.esriuk.com/server/rest/services/Hosted/Australia_Aug26_Peril5/FeatureServer/0",
    #     "name_field": "AccountName",
    #     "id_field":   "Locid"
    # },
    # "Philippines WS": {
    #     "url":        "https://axaxl-gis.cloud.esriuk.com/server/rest/services/Hosted/Philippines_May26_Peril2/FeatureServer/0",
    #     "name_field": "AccountName",
    #     "id_field":   "Locid"
    # },
}

EVENT_LAYERS = {
    # "Wildfire Perimeters": {
    #     "url":        "https://axaxl-gis.cloud.esriuk.com/server/rest/services/Hosted/Wildfire_Perimeters/FeatureServer/0",
    #     "name_field": "IncidentName",
    #     "id_field":   "GeometryID"
    # },
    "Hurricane Tracks": {
        "url":        "https://services9.arcgis.com/RHVPKKiFTONKtxq3/ArcGIS/rest/services/Recent_Hurricanes_v1/FeatureServer/2",
        "name_field": "StormName",
        "id_field":   "StormID"
    },
    "Flood Events": {
        "url":        "https://axaxl-gis.cloud.esriuk.com/server/rest/services/Hosted/ICEYE_Texas_Flood_July_26_Extent/FeatureServer/0",
        "name_field": "flooded_cities",
        "id_field":   "flood_id"
    },
}

GLOBAL_THRESHOLD        = 1
REGION_THRESHOLDS       = {
    "Florida WS": 1,
}

ALERT_MODE              = "BOTH"
DIGEST_HOUR             = 7
MAX_PARALLEL_WORKERS    = 8
SUSPECT_COUNT_THRESHOLD = 15000

# --- SharePoint Configuration ---
SHAREPOINT_SITE_URL     = "https://xlgroupcrmeur.sharepoint.com/:l:/r/teams/GIS90/Lists/Nebula%20Alerts?e=WLRnxL"
SHAREPOINT_LIST_NAME    = "Nebula Alerts"
SHAREPOINT_USERNAME     = "leo.hussain@axaxl.com"
SHAREPOINT_PASSWORD     = 'Khaleesi12;;'#os.environ.get("SHAREPOINT_PASSWORD", "")

# --- State Files ---
ALERT_STATE_FILE        = "nebula_alert_state.json"
DIGEST_STATE_FILE       = "nebula_digest_state.json"

# =============================================================================
# STATE MANAGEMENT
# =============================================================================

def load_json(filepath):
    if os.path.exists(filepath):
        with open(filepath, "r") as f:
            return 
