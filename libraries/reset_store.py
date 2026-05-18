# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details


def _fetch_workspaces(api, location_id: str) -> list:
    """Return all workspaces at the given location."""
    result = api.call(
        'GET', 'workspaces',
        params={'orgId': api.org_id, 'locationId': location_id, 'max': 1000}
    )
    if 'error' in result:
        return []
    return result.get('items', [])


def _fetch_hunt_groups(api, location_id: str) -> list:
    """Return all hunt groups at the given location."""
    result = api.call(
        'GET', 'telephony/config/huntGroups',
        params={'orgId': api.org_id, 'locationId': location_id}
    )
    if 'error' in result:
        return []
    return result.get('huntGroups', [])


def _fetch_auto_attendants(api, location_id: str) -> list:
    """Return all auto attendants at the given location."""
    result = api.call(
        'GET', 'telephony/config/autoAttendants',
        params={'orgId': api.org_id, 'locationId': location_id}
    )
    if 'error' in result:
        return []
    return result.get('autoAttendants', [])


def _fetch_location_announcements(api, location_id: str) -> list:
    """Return all announcement files at the given location (level=LOCATION)."""
    result = api.call(
        'GET', 'telephony/config/announcements',
        params={'orgId': api.org_id, 'locationId': location_id}
    )
    if 'error' in result:
        return []
    # Filter to only location-level files (exclude org-level files)
    return [
        a for a in result.get('announcements', [])
        if str(a.get('level', '')).upper() == 'LOCATION'
    ]


def _fetch_call_park_groups(api, location_id: str) -> list:
    """Return all call park groups at the given location."""
    result = api.call(
        'GET', f'telephony/config/locations/{location_id}/callParks',
        params={'orgId': api.org_id}
    )
    if 'error' in result:
        return []
    return result.get('callParks', [])


def _fetch_call_park_extensions(api, location_id: str) -> list:
    """Return all call park extensions at the given location."""
    result = api.call(
        'GET', 'telephony/config/callParkExtensions',
        params={'orgId': api.org_id, 'locationId': location_id}
    )
    if 'error' in result:
        return []
    return result.get('callParkExtensions', [])


def reset_store(api, location_data: dict, filepath: str, read_excel_sheet):
    """
    Reset a store location back to a clean state so it can be re-imported.

    Deletes (in order):
      1. All workspaces at the location  (devices are auto-removed with their workspace)
      2. All hunt groups at the location
      3. All auto attendants at the location
      4. All call park extensions at the location
    """
    location_id   = location_data['id']
    location_name = location_data['name']

    print(f"\n{'='*60}")
    print("Reset Store — Gathering inventory")
    print(f"{'='*60}")

    # -----------------------------------------------------------------------
    # 1. Collect everything that will be deleted
    # -----------------------------------------------------------------------
    print(f"\n  Fetching workspaces...")
    workspaces = _fetch_workspaces(api, location_id)

    print(f"  Fetching hunt groups...")
    hunt_groups = _fetch_hunt_groups(api, location_id)

    print(f"  Fetching auto attendants...")
    auto_attendants = _fetch_auto_attendants(api, location_id)

    print(f"  Fetching call park extensions...")
    call_park_exts = _fetch_call_park_extensions(api, location_id)

    print(f"  Fetching call park groups...")
    call_park_groups = _fetch_call_park_groups(api, location_id)

    print(f"  Fetching location announcements...")
    announcements = _fetch_location_announcements(api, location_id)

    # -----------------------------------------------------------------------
    # 2. Display full deletion manifest
    # -----------------------------------------------------------------------
    print(f"\n{'='*60}")
    print(f"RESET STORE — Deletion Manifest for: {location_name}")
    print(f"{'='*60}")

    print(f"\n  Workspaces to delete ({len(workspaces)}):")
    if workspaces:
        for ws in workspaces:
            print(f"    - {ws.get('displayName', ws.get('id'))}")
    else:
        print(f"    (none found)")

    print(f"\n  Hunt Groups to delete ({len(hunt_groups)}):")
    if hunt_groups:
        for hg in hunt_groups:
            print(f"    - {hg.get('name', hg.get('id'))}")
    else:
        print(f"    (none found)")

    print(f"\n  Auto Attendants to delete ({len(auto_attendants)}):")
    if auto_attendants:
        for aa in auto_attendants:
            print(f"    - {aa.get('name', aa.get('id'))}")
    else:
        print(f"    (none found)")

    print(f"\n  Call Park Groups to delete ({len(call_park_groups)}):")
    if call_park_groups:
        for cpg in call_park_groups:
            print(f"    - {cpg.get('name', cpg.get('id'))}")
    else:
        print(f"    (none found)")

    print(f"\n  Call Park Extensions to delete ({len(call_park_exts)}):")
    if call_park_exts:
        for cpe in call_park_exts:
            print(f"    - {cpe.get('name', cpe.get('id'))}  (ext: {cpe.get('extension', '')})")
    else:
        print(f"    (none found)")

    print(f"\n  Location Announcements to delete ({len(announcements)}):")
    if announcements:
        for ann in announcements:
            print(f"    - {ann.get('fileName', ann.get('id'))}")
    else:
        print(f"    (none found)")

    print(f"\n  Note: Devices are automatically removed when their workspace is deleted.")

    total = len(workspaces) + len(hunt_groups) + len(auto_attendants) + len(call_park_groups) + len(call_park_exts) + len(announcements)
    if total == 0:
        print(f"\n  Nothing to delete. Location is already clean.")
        return

    # -----------------------------------------------------------------------
    # 3. Confirm by typing location name in ALL CAPS
    # -----------------------------------------------------------------------
    print(f"\n{'='*60}")
    print(f"  WARNING: This will permanently delete {total} item(s) from Control Hub.")
    print(f"  This action cannot be undone.")
    print(f"\n  To confirm, type the location name in ALL CAPS: {location_name.upper()}")
    print(f"{'='*60}")
    confirmation = input("  Confirmation: ").strip()

    if confirmation != location_name.upper():
        print(f"\n  Confirmation did not match. Reset Store cancelled.")
        return

    # -----------------------------------------------------------------------
    # 4. Execute deletions
    # -----------------------------------------------------------------------
    results = {'deleted': 0, 'failed': 0, 'errors': []}

    def _delete(label: str, endpoint: str, item_name: str):
        result = api.call('DELETE', endpoint, params={'orgId': api.org_id})
        if 'error' in result:
            print(f"  FAILED  [{label}] {item_name}: {result['error']}")
            results['failed'] += 1
            results['errors'].append(f"{label} '{item_name}': {result['error']}")
        else:
            print(f"  Deleted [{label}] {item_name}")
            results['deleted'] += 1

    # --- Workspaces (devices cascade-deleted automatically) ---
    if workspaces:
        print(f"\n{'='*60}")
        print("Deleting Workspaces")
        print(f"{'='*60}")
        for ws in workspaces:
            _delete('Workspace', f"workspaces/{ws['id']}", ws.get('displayName', ws['id']))

    # --- Hunt Groups ---
    if hunt_groups:
        print(f"\n{'='*60}")
        print("Deleting Hunt Groups")
        print(f"{'='*60}")
        for hg in hunt_groups:
            _delete('Hunt Group',
                    f"telephony/config/locations/{location_id}/huntGroups/{hg['id']}",
                    hg.get('name', hg['id']))

    # --- Auto Attendants ---
    if auto_attendants:
        print(f"\n{'='*60}")
        print("Deleting Auto Attendants")
        print(f"{'='*60}")
        for aa in auto_attendants:
            _delete('Auto Attendant',
                    f"telephony/config/locations/{location_id}/autoAttendants/{aa['id']}",
                    aa.get('name', aa['id']))

    # --- Call Park Groups (must be deleted before extensions) ---
    if call_park_groups:
        print(f"\n{'='*60}")
        print("Deleting Call Park Groups")
        print(f"{'='*60}")
        for cpg in call_park_groups:
            _delete('Call Park Group',
                    f"telephony/config/locations/{location_id}/callParks/{cpg['id']}",
                    cpg.get('name', cpg['id']))

    # --- Call Park Extensions ---
    if call_park_exts:
        print(f"\n{'='*60}")
        print("Deleting Call Park Extensions")
        print(f"{'='*60}")
        for cpe in call_park_exts:
            _delete('Call Park Ext',
                    f"telephony/config/locations/{location_id}/callParkExtensions/{cpe['id']}",
                    cpe.get('name', cpe['id']))

    # --- Location Announcements ---
    if announcements:
        print(f"\n{'='*60}")
        print("Deleting Location Announcements")
        print(f"{'='*60}")
        for ann in announcements:
            _delete('Announcement',
                    f"telephony/config/locations/{location_id}/announcements/{ann['id']}",
                    ann.get('fileName', ann['id']))

    # -----------------------------------------------------------------------
    # 5. Summary
    # -----------------------------------------------------------------------
    print(f"\n{'='*60}")
    print("Reset Store Summary")
    print(f"{'='*60}")
    print(f"  Successfully deleted: {results['deleted']}")
    print(f"  Failed:               {results['failed']}")
    if results['errors']:
        print(f"\n  Errors:")
        for e in results['errors']:
            print(f"    - {e}")
    print(f"{'='*60}")
