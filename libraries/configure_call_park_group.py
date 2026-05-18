# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

# Sheet layout (0-indexed, first sheet / location tab)
# Row 48 (idx 47): "Call Park Group"  header
# Row 49 (idx 48): "Calling - Features - Call Park Group"
# Row 50 (idx 49): "Location",  col B = location name
# Row 51 (idx 50): "Call Park Group Name", col B = group name
# Row 52 (idx 51): "Members",   col B = description (all workspaces except Paging)
# Row 53 (idx 52): "Park Destinations", col B = description (all park extensions)

ROW_GROUP_NAME   = 50   # 0-indexed row for group name (row 51 in sheet)
COL_VALUE        = 1    # column B


def _read_call_park_group_config(sheet_data: list) -> dict | None:
    """
    Parse call park group config from the location tab.
    Returns dict with 'name', or None if data is missing.
    """
    if len(sheet_data) <= ROW_GROUP_NAME:
        return None

    row = sheet_data[ROW_GROUP_NAME]
    name = str(row[COL_VALUE]).strip() if len(row) > COL_VALUE and row[COL_VALUE] else ""

    if not name:
        return None

    return {'name': name}


def _fetch_workspaces(api, location_id: str) -> list:
    """Return all workspaces at the location."""
    result = api.call(
        'GET', 'workspaces',
        params={'orgId': api.org_id, 'locationId': location_id, 'max': 1000}
    )
    if 'error' in result:
        return []
    return result.get('items', [])


def _fetch_call_park_extensions(api, location_id: str) -> list:
    """Return all call park extensions at the location."""
    result = api.call(
        'GET', 'telephony/config/callParkExtensions',
        params={'orgId': api.org_id, 'locationId': location_id}
    )
    if 'error' in result:
        return []
    return result.get('callParkExtensions', [])


def _fetch_existing_call_park_group(api, location_id: str, name: str) -> str | None:
    """Return the ID of an existing call park group with the given name, or None."""
    result = api.call(
        'GET', f'telephony/config/locations/{location_id}/callParks',
        params={'orgId': api.org_id, 'name': name}
    )
    if 'error' in result:
        return None
    for cp in result.get('callParks', []):
        if cp.get('name', '').strip().lower() == name.strip().lower():
            return cp.get('id')
    return None


def configure_call_park_group(api, location_data: dict, filepath: str, read_excel_sheet):
    """
    Read call park group config from the location tab and create/update it in Control Hub.

    Members  : all workspaces at the location, excluding any named "Paging"
    Destinations: all call park extensions at the location
    """
    print(f"\n{'='*60}")
    proceed = input("\nProceed with Call Park Group configuration? (Y/n): ").strip().lower()
    if proceed not in ['', 'y', 'yes']:
        print("\nCall Park Group configuration skipped.")
        return

    location_id   = location_data['id']
    location_name = location_data['name']

    # -----------------------------------------------------------------------
    # 1. Read config from sheet
    # -----------------------------------------------------------------------
    sheet_data = read_excel_sheet(filepath, location_name)
    if not sheet_data:
        print(f"  Error: Could not read location tab '{location_name}'.")
        return

    config = _read_call_park_group_config(sheet_data)
    if not config:
        print(f"  Error: Call Park Group name not found in location tab (row 51, col B).")
        return

    group_name = config['name']
    print(f"\n  Call Park Group Name: {group_name}")

    # -----------------------------------------------------------------------
    # 2. Resolve members — all workspaces except those named "Paging"
    # -----------------------------------------------------------------------
    print(f"  Fetching workspaces at location...")
    workspaces = _fetch_workspaces(api, location_id)
    members = [
        ws['id']
        for ws in workspaces
        if 'paging' not in ws.get('displayName', '').lower()
    ]
    excluded = [ws.get('displayName') for ws in workspaces
                if 'paging' in ws.get('displayName', '').lower()]

    print(f"  Members: {len(members)} workspace(s)")
    if excluded:
        print(f"  Excluded (Paging): {', '.join(excluded)}")

    if not members:
        print(f"  Error: No eligible member workspaces found. Ensure workspaces have been imported first.")
        return

    # -----------------------------------------------------------------------
    # 3. Resolve park destinations — all call park extensions at the location
    # -----------------------------------------------------------------------
    print(f"  Fetching call park extensions...")
    cpe_list = _fetch_call_park_extensions(api, location_id)
    park_destinations = [cpe['id'] for cpe in cpe_list]

    print(f"  Park Destinations: {len(park_destinations)} extension(s)")
    if not park_destinations:
        print(f"  Warning: No call park extensions found. Ensure call park extensions have been created first.")

    # -----------------------------------------------------------------------
    # 4. Display configuration summary
    # -----------------------------------------------------------------------
    print(f"\n  Configuration Summary:")
    print(f"  {'Name':<30} {group_name}")
    print(f"  {'Members':<30} {len(members)} workspace(s)")
    print(f"  {'Park Destinations':<30} {len(park_destinations)} extension(s)")
    print(f"  {'Park on Agents Enabled':<30} False")

    # -----------------------------------------------------------------------
    # 5. Check for existing group with same name
    # -----------------------------------------------------------------------
    existing_id = _fetch_existing_call_park_group(api, location_id, group_name)

    payload = {
        'name':                 group_name,
        'recall':               {'option': 'ALERT_PARKING_USER_ONLY'},
        'agents':               [ws['id'] for ws in workspaces if 'paging' not in ws.get('displayName', '').lower()],
        'callParkExtensions':   [cpe['id'] for cpe in cpe_list],
        'parkOnAgentsEnabled':  False,
    }

    if existing_id:
        print(f"\n  *** A Call Park Group named '{group_name}' already exists ***")
        print(f"  Options:")
        print(f"    U - Update existing group with current members and destinations")
        print(f"    S - Skip")
        choice = input("  Choice (U/S): ").strip().upper()
        if choice != 'U':
            print(f"  Skipped: '{group_name}' left unchanged.")
            return

        print(f"\n  Updating Call Park Group...")
        result = api.call(
            'PUT',
            f'telephony/config/locations/{location_id}/callParks/{existing_id}',
            data=payload,
            params={'orgId': api.org_id}
        )
        if 'error' in result:
            print(f"  Error: {result['error']}")
        else:
            print(f"  Updated successfully.")
    else:
        print(f"\n  Creating Call Park Group...")
        result = api.call(
            'POST',
            f'telephony/config/locations/{location_id}/callParks',
            data=payload,
            params={'orgId': api.org_id}
        )
        if 'error' in result:
            print(f"  Error: {result['error']}")
        else:
            group_id = result.get('id')
            print(f"  Created successfully (ID: {group_id})")
