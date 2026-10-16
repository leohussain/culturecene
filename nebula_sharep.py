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
