#!/usr/bin/env python3
"""Complete ArcGIS -> existing Nebula Alerts SharePoint list script.

Install in your ArcGIS-compatible Python environment:
    python -m pip install arcgis "Office365-REST-Python-Client>=3.2,<4" filelock

Run once:            python nebula_sharepoint.py
Run every 30 mins:   python nebula_sharepoint.py --watch
Analyse only:       python nebula_sharepoint.py --dry-run
Test list login:    python nebula_sharepoint.py --check-sharepoint

Passwords: ARCGIS_PASSWORD and SHAREPOINT_PASSWORD environment variables,
or one prompt per password on interactive launch. No password temp files.
For an unattended run, supply the variables through your approved secret store.

SharePoint Online username/password login in the current Office365 library
requires SHAREPOINT_TENANT_ID and SHAREPOINT_CLIENT_ID as well. The app must
have approved delegated SharePoint permissions and the user must have list write
access. Password login cannot satisfy an MFA requirement. No Graph API is used.
Library authentication reference:
https://github.com/vgrem/office365-rest-python-client

No extra SharePoint columns are required. Rows are matched using your existing
Title + EventID + Region fields. Use one running instance/host. Without a unique
list constraint this is not a distributed, exactly-once delivery mechanism.

Use the actual internal SharePoint field names in SHAREPOINT_FIELDS if they
differ from the display names. EventID must be Text; Count must allow blanks for ERROR/TIMEOUT.
AffectedLocations should be multiple lines of plain text for long lists; with
a single-line Text column the script automatically limits it to that capacity.

This version preserves your active layers and ArcGIS verify_cert=False setting.
The July flood extent remains a static test unless the source itself is updated.
The hurricane layer is observed wind swath, not a forecast track. Event filters
default to 1=1, as in the original; set each 'where' to your live-event rule.
For trusted TLS, set ARCGIS_VERIFY_CERT=True and configure your approved CA.

The script updates the current row every check, including CLEAN results. It
keeps pending list writes in the JSON state for retry. It does not send email
or Teams messages, so the old ALERT_MODE/digest settings are not needed.
State/results may contain confidential location names; protect their directory.
Run as a .py file, not as pasted multiprocessing code inside a notebook.
"""
import argparse
import getpass
import hashlib
import json
import multiprocessing
import os
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

import requests
from filelock import FileLock, Timeout as LockTimeout

# =============================================================================
# CONFIGURATION — your layers and existing SharePoint list
# =============================================================================
PORTAL_URL = "https://axaxl-gis.cloud.esriuk.com/portal"
USERNAME = os.environ.get("ARCGIS_USERNAME", "builtin.admin")
ARCGIS_VERIFY_CERT = False  # Preserved from your original configuration.

PORTFOLIO_LAYERS = {
    "Florida WS": {
        "url": "https://axaxl-gis.cloud.esriuk.com/server/rest/services/Hosted/Florida_Aug26_Peril2/FeatureServer/0",
        "name_field": "AccountName", "id_field": "Locid",
    },
    # "Australia WF": {
    #     "url": "https://axaxl-gis.cloud.esriuk.com/server/rest/services/Hosted/Australia_Aug26_Peril5/FeatureServer/0",
    #     "name_field": "AccountName", "id_field": "Locid",
    # },
    # "Philippines WS": {
    #     "url": "https://axaxl-gis.cloud.esriuk.com/server/rest/services/Hosted/Philippines_May26_Peril2/FeatureServer/0",
    #     "name_field": "AccountName", "id_field": "Locid",
    # },
}

EVENT_LAYERS = {
    "Hurricane Tracks": {  # Display title retained; this URL is Observed Wind Swath.
        "url": "https://services9.arcgis.com/RHVPKKiFTONKtxq3/ArcGIS/rest/services/Recent_Hurricanes_v1/FeatureServer/2",
        "name_field": "STORMNAME", "id_field": "STORMID", "public": True,
        "where": "1=1",
    },
    "Flood Events": {
        "url": "https://axaxl-gis.cloud.esriuk.com/server/rest/services/Hosted/ICEYE_Texas_Flood_July_26_Extent/FeatureServer/0",
        "name_field": "flooded_cities", "id_field": "flood_id", "where": "1=1",
    },
    # "Wildfire Perimeters": {
    #     "url": "https://axaxl-gis.cloud.esriuk.com/server/rest/services/Hosted/Wildfire_Perimeters/FeatureServer/0",
    #     "name_field": "IncidentName", "id_field": "GeometryID", "where": "1=1",
    # },
}

GLOBAL_THRESHOLD = 1
REGION_THRESHOLDS = {"Florida WS": 1}
MAX_PARALLEL_WORKERS = 8
SUSPECT_COUNT_THRESHOLD = 15000
ANALYSIS_TIMEOUT_SECONDS = 1200  # Whole pool, including queued tasks.
CHECK_INTERVAL_SECONDS = 30 * 60
MAX_AFFECTED_LOCATIONS = 500    # Count is always the full count; names are a preview.

# Use the SITE URL, not the /:l:/r/ sharing link.
SHAREPOINT_SITE_URL = "https://xlgroupcrmeur.sharepoint.com/teams/GIS90"
SHAREPOINT_LIST_NAME = "Nebula Alerts"
SHAREPOINT_USERNAME = os.environ.get("SHAREPOINT_USERNAME", "leo.hussain@axaxl.com")
SHAREPOINT_TENANT_ID = os.environ.get("SHAREPOINT_TENANT_ID", "")
SHAREPOINT_CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID", "")
SHAREPOINT_FIELDS = {name: name for name in (
    "Title", "EventName", "EventID", "Region", "Count", "Threshold", "Status",
    "AlertReason", "AffectedLocations", "Timestamp")}

BASE_DIR = Path(__file__).resolve().parent
ALERT_STATE_FILE = BASE_DIR / "nebula_alert_state.json"
RESULTS_FILE = BASE_DIR / "nebula_results.json"
_WORKER_PASSWORD = None
_WORKER_GIS = None


class MonitorError(RuntimeError):
    """Application validation error safe to show without raw HTTP response data."""
    pass


def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


def password(name, prompt):
    value = os.environ.get(name)
    if value:
        return value
    if not sys.stdin.isatty():
        raise MonitorError(f"Set {name} for unattended execution")
    return getpass.getpass(prompt)


def load_json(filepath):
    if Path(filepath).exists():
        with open(filepath, "r", encoding="utf-8") as f:
            return json.load(f)  # A damaged state file must not silently reset alerts.
    return {"version": "nebula-sharepoint-1", "records": {}, "pending": []}


def save_json(filepath, data):
    fd, temporary = tempfile.mkstemp(dir=Path(filepath).parent, prefix="nebula_state_")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False, allow_nan=False)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temporary, filepath)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)


def key_for(event_type, event_id, region):
    return json.dumps([event_type, str(event_id), region], separators=(",", ":"))


def safe_error(exc):
    # Do not print exception URLs/bodies, which can contain tokens or account data.
    response = getattr(exc, "response", None)
    code = getattr(response, "status_code", None)
    if isinstance(exc, (requests.exceptions.Timeout, TimeoutError)):
        return "TIMEOUT", "Network request timed out"
    if isinstance(exc, requests.exceptions.SSLError):
        return "ERROR", "TLS certificate validation failed"
    if isinstance(exc, requests.exceptions.ProxyError):
        return "ERROR", "Proxy connection failed"
    if isinstance(exc, MonitorError):
        return "ERROR", str(exc)  # Our own validation errors only; see callers.
    return "ERROR", f"{type(exc).__name__}" + (f" (HTTP {code})" if code else "")


def connect_gis(secret):
    from arcgis.gis import GIS
    return GIS(PORTAL_URL, USERNAME, secret, verify_cert=ARCGIS_VERIFY_CERT,
               set_active=False)


def initialise_worker(secret):
    global _WORKER_PASSWORD, _WORKER_GIS
    # Memory only; lazy connection inside the task lets failures return normally.
    _WORKER_PASSWORD, _WORKER_GIS = secret, None


def validate_layer(layer, config, expected_geometry):
    props = layer.properties
    if props.geometryType != expected_geometry:
        raise MonitorError(f"Expected {expected_geometry}; received {props.geometryType}")
    names = {f["name"] for f in props.fields}
    for field in (config["name_field"], config["id_field"]):
        if field not in names:
            raise MonitorError(f"Layer field does not exist: {field}")


def query_ids(layer, **kwargs):
    result = layer.query(return_ids_only=True, **kwargs)
    if result.get("exceededTransferLimit") or not isinstance(result.get("objectIds"), list):
        raise MonitorError("Incomplete ArcGIS ID response")
    ids = set(result["objectIds"])
    if layer.query(return_count_only=True, **kwargs) != len(ids):
        raise MonitorError("ArcGIS count/ID mismatch; source changed or response was truncated")
    return ids


def get_features(layer, ids, fields, geometry=False):
    oid = layer.properties.objectIdField
    size = max(1, min(200, layer.properties.get("maxRecordCount", 200)))
    ids, features = sorted(ids), []
    for start in range(0, len(ids), size):
        batch = ids[start:start + size]
        found = layer.query(object_ids=",".join(map(str, batch)),
                            out_fields=",".join(sorted(set(fields) | {oid})),
                            return_geometry=geometry, out_sr=4326,
                            return_all_records=True)
        if (len(found.features) != len(batch) or
                {f.attributes[oid] for f in found.features} != set(batch)):
            raise MonitorError("Incomplete ArcGIS feature response; will retry next scan")
        features.extend(found.features)
    return features


def fetch_events(event_type, config, gis):
    from arcgis.features import FeatureLayer
    from arcgis.gis import GIS
    # Never send the Enterprise login to the public event service.
    event_gis = GIS(set_active=False) if config.get("public") else gis
    layer = FeatureLayer(config["url"], gis=event_gis)
    validate_layer(layer, config, "esriGeometryPolygon")
    ids = query_ids(layer, where=config.get("where", "1=1"))
    features = get_features(layer, ids, [config["id_field"], config["name_field"]], True)
    grouped = {}
    for feature in features:
        value = feature.attributes.get(config["id_field"])
        if value is None or not str(value).strip() or str(value) == "__SOURCE__":
            raise MonitorError("A source event ID is blank or uses reserved value __SOURCE__")
        event_id = str(value).strip()
        shape = dict(feature.geometry or {})
        if not shape.get("rings"):
            raise MonitorError("Event feature is missing polygon rings")
        shape.setdefault("spatialReference", {"wkid": 4326})
        sr = shape["spatialReference"]
        if sr.get("latestWkid", sr.get("wkid")) != 4326:
            raise MonitorError("Event geometry was not returned in EPSG:4326")
        item = grouped.setdefault(event_id, {"EventID": event_id,
            "EventName": str(feature.attributes.get(config["name_field"]) or event_id),
            "polygons": []})
        item["polygons"].append(shape)
    return grouped


def analyse_one(task):
    from arcgis.features import FeatureLayer
    from arcgis.geometry import Geometry
    from arcgis.geometry.filters import intersects
    global _WORKER_GIS
    try:
        if _WORKER_GIS is None:
            _WORKER_GIS = connect_gis(_WORKER_PASSWORD)
        config = PORTFOLIO_LAYERS[task["Region"]]
        layer = FeatureLayer(config["url"], gis=_WORKER_GIS)
        validate_layer(layer, config, "esriGeometryPoint")
        ids = set()
        for polygon in task["polygons"]:
            ids |= query_ids(layer, where=config.get("where", "1=1"),
                             geometry_filter=intersects(Geometry(polygon), sr=4326))
        selected = sorted(ids)[:MAX_AFFECTED_LOCATIONS]
        features = get_features(layer, selected, [config["name_field"], config["id_field"]])
        labels = []
        for feature in features:
            attrs = feature.attributes
            location = attrs.get(config["id_field"], attrs[layer.properties.objectIdField])
            labels.append(f"{attrs.get(config['name_field']) or 'Unnamed'} [{location}]")
        names = ", ".join(sorted(labels))
        if len(ids) > len(selected):
            names += f" ... [{len(ids) - len(selected)} more locations; full count retained]"
        fingerprint = hashlib.sha256(json.dumps(sorted(ids)).encode()).hexdigest()
        return {"Count": len(ids), "AffectedLocations": names, "fingerprint": fingerprint}
    except Exception as exc:
        status, reason = safe_error(exc)
        return {"Count": None, "Status": status, "AlertReason": reason}


def run_parallel(tasks, secret):
    if not tasks:
        return []
    # spawn is safe on Windows when launched from this guarded .py entry point.
    pool = multiprocessing.get_context("spawn").Pool(
        min(MAX_PARALLEL_WORKERS, len(tasks)), initializer=initialise_worker, initargs=(secret,))
    pending = [(task, pool.apply_async(analyse_one, (task,))) for task in tasks]
    deadline, results = time.monotonic() + ANALYSIS_TIMEOUT_SECONDS, []
    try:
        while pending:
            remaining = []
            for task, future in pending:
                if future.ready():
                    try:
                        outcome = future.get()
                    except Exception as exc:
                        status, reason = safe_error(exc)
                        outcome = {"Count": None, "Status": status, "AlertReason": reason}
                    results.append((task, outcome))
                else:
                    remaining.append((task, future))
            pending = remaining
            if pending and time.monotonic() >= deadline:
                results.extend((task, {"Count": None, "Status": "TIMEOUT",
                    "AlertReason": "Analysis run deadline reached; result unknown"}) for task, _ in pending)
                break
            if pending:
                time.sleep(0.25)
    finally:
        # Also terminates outstanding workers after the pool deadline.
        pool.terminate()
        pool.join()
    return results


def make_record(task, outcome, previous):
    threshold = REGION_THRESHOLDS.get(task["Region"], GLOBAL_THRESHOLD)
    count = outcome.get("Count")
    active = previous.get("active", False)
    if outcome.get("Status") in ("ERROR", "TIMEOUT"):
        status, reason = outcome["Status"], outcome["AlertReason"]
    elif count > SUSPECT_COUNT_THRESHOLD:
        status, reason = "SUSPECT", "Count exceeds the sanity threshold; investigate geometry/data"
    elif count >= threshold:
        status = "BREACH"
        if not active:
            reason = "New breach"
        elif previous.get("last_good_count") != count:
            reason = f"Count changed from {previous.get('last_good_count')} to {count}"
        elif previous.get("fingerprint") != outcome.get("fingerprint"):
            reason = "Affected locations changed; count is unchanged"
        else:
            reason = "Breach remains active"
        active = True
    else:
        status = "ALL CLEAR" if active or previous.get("fields", {}).get("Status") == "ALL CLEAR" else "CLEAN"
        reason = "Previous breach is now below threshold" if status == "ALL CLEAR" else "Count is below threshold"
        active = False
    fields = {name: task[name] for name in ("Title", "EventName", "EventID", "Region")}
    fields.update(Count=count, Threshold=threshold, Status=status, AlertReason=reason,
                  AffectedLocations=outcome.get("AffectedLocations", ""), Timestamp=now())
    record = {"fields": fields, "active": active,
              "last_good_count": previous.get("last_good_count"),
              "fingerprint": previous.get("fingerprint")}
    if status in ("BREACH", "CLEAN", "ALL CLEAR"):
        record.update(last_good_count=count, fingerprint=outcome.get("fingerprint"))
    return record


# =============================================================================
# SHAREPOINT — writes your existing columns; no new AlertKey column
# =============================================================================
def connect_sharepoint(secret):
    from office365.sharepoint.client_context import ClientContext
    if not SHAREPOINT_TENANT_ID or not SHAREPOINT_CLIENT_ID:
        raise MonitorError("Set SHAREPOINT_TENANT_ID and SHAREPOINT_CLIENT_ID. Current Office365 "
                           "username/password login needs both; legacy UserCredential login is retired.")
    ctx = ClientContext(SHAREPOINT_SITE_URL).with_username_and_password(
        SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_USERNAME, secret)
    def set_timeout(request):
        request.timeout = 60
    ctx.pending_request().beforeExecute += set_timeout
    target = ctx.web.lists.get_by_title(SHAREPOINT_LIST_NAME)
    target.get().execute_query()  # Fails here if authentication/list access is denied.
    columns = target.fields.get().execute_query()
    definitions = {f.properties["InternalName"]: f.properties for f in columns}
    for logical, internal in SHAREPOINT_FIELDS.items():
        if internal not in definitions:
            raise MonitorError(f"SharePoint internal field not found: {internal}; edit SHAREPOINT_FIELDS")
    if definitions[SHAREPOINT_FIELDS["Count"]].get("Required"):
        raise MonitorError("SharePoint Count must allow blanks for ERROR/TIMEOUT results")
    return target, definitions


def odata_string(value):
    return "'" + str(value).replace("'", "''") + "'"


def write_sharepoint(target, definitions, fields):
    # Identity uses the full values, never truncated strings.
    for name in ("Title", "EventID", "Region"):
        maximum = int(definitions[SHAREPOINT_FIELDS[name]].get("MaxLength") or 255)
        if len(str(fields[name])) > maximum:
            raise MonitorError(f"{name} exceeds the list key length; choose a shorter stable identifier")
    where = " and ".join(f"{SHAREPOINT_FIELDS[name]} eq {odata_string(fields[name])}"
                         for name in ("Title", "EventID", "Region"))
    existing = target.items.filter(where).top(2).get().execute_query()
    if len(existing) > 1:
        raise MonitorError("Duplicate Title/EventID/Region rows already exist in Nebula Alerts")
    payload = {}
    for logical, value in fields.items():
        internal = SHAREPOINT_FIELDS[logical]
        definition = definitions[internal]
        if isinstance(value, str) and definition.get("TypeAsString") in ("Text", "Note"):
            maximum = int(definition.get("MaxLength") or (255 if definition["TypeAsString"] == "Text" else 60000))
            if len(value) > maximum:
                value = value[:max(0, maximum - 15)] + " ...[truncated]"
        payload[internal] = value
    if existing:
        item = existing[0]
        for name, value in payload.items():
            item.set_property(name, value)
        item.update().execute_query()
    else:
        target.add_item(payload).execute_query()


def deliver_pending(state, target, definitions):
    # FIFO. A failed write remains pending and is searched for again before create.
    while state["pending"]:
        write_sharepoint(target, definitions, state["pending"][0])
        state["pending"].pop(0)
        save_json(ALERT_STATE_FILE, state)


def run_once(arcgis_secret, sharepoint_secret, dry_run=False):
    state = load_json(ALERT_STATE_FILE)
    if state.get("version") != "nebula-sharepoint-1":
        raise MonitorError("State has a different format. Back it up/rename it before starting this version.")
    tasks, results = [], []
    gis = connect_gis(arcgis_secret)
    for event_type, config in EVENT_LAYERS.items():
        try:
            events = fetch_events(event_type, config, gis)
            source_error = None
        except Exception as exc:
            events = {}
            status, reason = safe_error(exc)
            source_error = {"Status": status, "AlertReason": reason, "Count": None}
        for region in PORTFOLIO_LAYERS:
            for event in events.values():
                tasks.append({"Title": event_type, "Region": region, **event})
            old = [r for r in state["records"].values()
                   if r["fields"]["Title"] == event_type and r["fields"]["Region"] == region]
            for record in old:
                fields = record["fields"]
                if fields["EventID"] not in events and fields["EventID"] != "__SOURCE__":
                    results.append((fields, source_error or {"Count": None, "Status": "ERROR",
                        "AlertReason": "Event no longer appears in the feed; resolution is unconfirmed"}))
            source_key = key_for(event_type, "__SOURCE__", region)
            if source_error or not events or source_key in state["records"]:
                health = source_error or {"Count": 0, "AffectedLocations": "", "fingerprint": "source-ok"}
                task = {"Title": event_type, "EventID": "__SOURCE__",
                        "EventName": "[Source check]", "Region": region}
                results.append((task, health))
    results.extend(run_parallel(tasks, arcgis_secret))
    records = []
    for task, outcome in results:
        key = key_for(task["Title"], task["EventID"], task["Region"])
        record = make_record(task, outcome, state["records"].get(key, {}))
        if task["EventID"] == "__SOURCE__" and not outcome.get("Status"):
            record["fields"]["AlertReason"] = "Event feed query succeeded. Source check only, not exposure clearance."
        records.append(record)
        if not dry_run:
            state["records"][key] = record
            state["pending"].append(record["fields"].copy())
        f = record["fields"]
        print(f"{f['Title']} | {f['Region']} | {f['Status']} | Count={f['Count']}")
    save_json(RESULTS_FILE, {"checked_at": now(), "results": [r["fields"] for r in records]})
    if dry_run:
        print(f"Dry run saved to {RESULTS_FILE}; no list or state changes.")
        return
    save_json(ALERT_STATE_FILE, state)  # Save results before attempting SharePoint.
    target, definitions = connect_sharepoint(sharepoint_secret)
    deliver_pending(state, target, definitions)
    print(f"SharePoint updated: {SHAREPOINT_LIST_NAME} ({len(records)} results this scan).")


def main():
    parser = argparse.ArgumentParser(description="ArcGIS intersections to Nebula Alerts")
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--watch", action="store_true", help="Repeat every 30 minutes")
    modes.add_argument("--dry-run", action="store_true", help="Query ArcGIS without writing SharePoint")
    modes.add_argument("--check-sharepoint", action="store_true", help="Test list authentication and field names")
    args = parser.parse_args()
    try:
        sp_secret = None if args.dry_run else password("SHAREPOINT_PASSWORD", "Microsoft 365 password: ")
        if args.check_sharepoint:
            connect_sharepoint(sp_secret)
            print("SharePoint login/list/fields OK. No items written.")
            return 0
        arcgis_secret = password("ARCGIS_PASSWORD", "ArcGIS password: ")
        # Holding this OS-backed lock also protects the pending queue across scans.
        with FileLock(str(ALERT_STATE_FILE) + ".lock", timeout=0):
            while True:
                started = time.monotonic()
                try:
                    run_once(arcgis_secret, sp_secret, args.dry_run)
                except Exception as exc:
                    status, reason = safe_error(exc)
                    print(f"{status}: {reason}. Pending writes, if saved, will retry next run.", file=sys.stderr)
                    if not args.watch:
                        return 1
                if not args.watch:
                    return 0
                delay = max(1, CHECK_INTERVAL_SECONDS - (time.monotonic() - started))
                print(f"Next check in {int(delay)} seconds; Ctrl+C stops the monitor.")
                time.sleep(delay)
    except LockTimeout:
        print("Another instance is using this state file; exiting.", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("Monitor stopped.")
        return 0
    except Exception as exc:
        _, reason = safe_error(exc)
        print(reason, file=sys.stderr)
        return 1


if __name__ == "__main__":
    multiprocessing.freeze_support()
    raise SystemExit(main())
