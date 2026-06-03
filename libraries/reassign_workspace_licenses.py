# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

import time
import requests
import json


def _paginated_get(api, endpoint, params):
    """
    Fetch all items from a paginated Webex API endpoint.
    Handles Link header pagination since api.call() only returns JSON body.
    """
    url = f"{api.base_url}/{endpoint}"
    headers = {
        "Authorization": f"Bearer {api.token}",
        "Content-Type": "application/json"
    }

    all_items = []

    while url:
        api.api_logger.info(f"API Call: GET {url}")
        if params:
            api.api_logger.info(f"Params: {json.dumps(params)}")

        try:
            response = requests.get(url, headers=headers, params=params)
            api.api_logger.info(f"Response Status: {response.status_code}")

            if response.status_code != 200:
                api.api_logger.error(f"API Error: {response.status_code} - {response.text}")
                return None, response.text

            data = response.json()
            items = data.get("items", [])
            all_items.extend(items)

            # Check for next page via Link header
            link_header = response.headers.get("Link", "")
            url = None
            params = None  # params are embedded in the next URL
            if link_header and 'rel="next"' in link_header:
                # Parse: <https://...>; rel="next"
                parts = link_header.split(";")
                if parts:
                    next_url = parts[0].strip().strip("<>")
                    url = next_url

        except Exception as e:
            api.api_logger.error(f"Exception during paginated GET: {e}")
            return None, str(e)

    return all_items, None


def reassign_workspace_licenses(api):
    """
    Move Webex Calling - Workspaces licenses from one subscription to another
    by updating each workspace's license assignment.
    """
    print("\n" + "=" * 60)
    print("Reassign Workspace Calling Subscription")
    print("=" * 60)

    # Step 1: Fetch all licenses
    print("\nFetching organization licenses...")
    lic_result = api.call("GET", "licenses", params={"orgId": api.org_id})

    if "error" in lic_result:
        print(f"Error fetching licenses: {lic_result['error']}")
        return

    all_licenses = lic_result.get("items", [])
    if not all_licenses:
        print("No licenses found.")
        return

    # Filter to Webex Calling - Workspaces licenses only
    workspace_licenses = [
        lic for lic in all_licenses
        if "workspace" in lic.get("name", "").lower() and "calling" in lic.get("name", "").lower()
    ]

    if len(workspace_licenses) < 2:
        print(f"\nFound {len(workspace_licenses)} Workspace Calling license(s).")
        print("Need at least 2 subscriptions with Workspace Calling licenses to reassign.")
        return

    # Display workspace calling licenses
    print(f"\nFound {len(workspace_licenses)} Workspace Calling license(s):\n")
    print(f"  {'#':<4} {'Subscription':<55} {'Consumed':>10} {'Total':>10} {'Available':>10}")
    print(f"  {'-' * 89}")

    for i, lic in enumerate(workspace_licenses, 1):
        sub_id = lic.get("subscriptionId", "Unknown")
        name = lic.get("name", "N/A")
        consumed = lic.get("consumedUnits", 0)
        total = lic.get("totalUnits", 0)
        available = total - consumed
        print(f"  {i:<4} {name} [{sub_id}]")
        print(f"       {'':45} {consumed:>10} {total:>10} {available:>10}")

    # Step 2: User selects source and target
    print()
    source_choice = input("Select SOURCE subscription # (licenses to move FROM): ").strip()
    try:
        source_lic = workspace_licenses[int(source_choice) - 1]
    except (ValueError, IndexError):
        print("Invalid selection.")
        return

    target_choice = input("Select TARGET subscription # (licenses to move TO): ").strip()
    try:
        target_lic = workspace_licenses[int(target_choice) - 1]
    except (ValueError, IndexError):
        print("Invalid selection.")
        return

    if source_lic["id"] == target_lic["id"]:
        print("Source and target cannot be the same license.")
        return

    source_id = source_lic["id"]
    target_id = target_lic["id"]
    source_sub = source_lic.get("subscriptionId", "Unknown")
    target_sub = target_lic.get("subscriptionId", "Unknown")

    print(f"\n  Source: {source_lic['name']} [{source_sub}] ({source_lic.get('consumedUnits', 0)} consumed)")
    print(f"  Target: {target_lic['name']} [{target_sub}] ({target_lic.get('consumedUnits', 0)} consumed / {target_lic.get('totalUnits', 0)} total)")

    # Step 3: Fetch all workspaces with pagination
    print("\nFetching all workspaces (this may take a moment)...")
    workspaces, err = _paginated_get(api, "workspaces", {"orgId": api.org_id, "max": 1000})

    if err:
        print(f"Error fetching workspaces: {err}")
        return

    if not workspaces:
        print("No workspaces found in the organization.")
        return

    print(f"  Retrieved {len(workspaces)} workspace(s) total.")

    # Step 4: Filter workspaces that have the source license
    # Licenses are nested at: calling.webexCalling.licenses[]
    affected_workspaces = []
    for ws in workspaces:
        calling = ws.get("calling", {})
        webex_calling = calling.get("webexCalling", {})
        ws_licenses = webex_calling.get("licenses", [])
        if source_id in ws_licenses:
            affected_workspaces.append(ws)

    if not affected_workspaces:
        print(f"\nNo workspaces found using the source license from {source_sub}.")
        return

    print(f"  Found {len(affected_workspaces)} workspace(s) assigned to source subscription [{source_sub}].")

    # Step 5: List affected workspaces and offer filter by selection
    # Show full list
    def _print_workspace_list(ws_list):
        print(f"\n  Workspaces to be reassigned ({len(ws_list)}):")
        print(f"  {'#':<6} {'Display Name'}")
        print(f"  {'-' * 60}")
        for idx, ws in enumerate(ws_list, 1):
            print(f"  {idx:<6} {ws.get('displayName', 'N/A')}")

    _print_workspace_list(affected_workspaces)

    # Offer number range filter
    print(f"\n  Filter by selection? This will limit the scope of workspaces to reassign.")
    print(f"  Enter a range (e.g. 1-4), comma-separated numbers (e.g. 1,3,5),")
    print(f"  a combination (e.g. 1-4,7,10-12), or press Enter to include all.")
    filter_input = input("  Selection: ").strip()

    if filter_input:
        # Parse selection input
        selected_indices = set()
        try:
            for part in filter_input.split(","):
                part = part.strip()
                if "-" in part:
                    start, end = part.split("-", 1)
                    start = int(start.strip())
                    end = int(end.strip())
                    for n in range(start, end + 1):
                        if 1 <= n <= len(affected_workspaces):
                            selected_indices.add(n - 1)
                else:
                    n = int(part)
                    if 1 <= n <= len(affected_workspaces):
                        selected_indices.add(n - 1)
        except ValueError:
            print("Invalid input. Aborted.")
            return

        if not selected_indices:
            print("No valid selections. Aborted.")
            return

        # Apply filter
        selected_indices = sorted(selected_indices)
        affected_workspaces = [affected_workspaces[i] for i in selected_indices]

        # Show filtered list
        _print_workspace_list(affected_workspaces)

    # Check target capacity against final workspace count
    target_available = target_lic.get("totalUnits", 0) - target_lic.get("consumedUnits", 0)
    if target_available < len(affected_workspaces):
        print(f"\n  WARNING: Target subscription has {target_available} available licenses")
        print(f"           but {len(affected_workspaces)} workspaces need to be moved.")
        proceed = input("  Continue anyway? (y/n): ").strip().lower()
        if proceed != 'y':
            print("Aborted.")
            return

    print(f"\n{'=' * 60}")
    print(f"  SUMMARY")
    print(f"{'=' * 60}")
    print(f"  Action: Move {len(affected_workspaces)} workspace(s)")
    print(f"  From:   {source_sub}")
    print(f"  To:     {target_sub}")
    print(f"{'=' * 60}")

    confirm = input("\nProceed with reassignment? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("Aborted.")
        return

    # Step 6: Execute reassignment
    print(f"\nReassigning {len(affected_workspaces)} workspace(s)...\n")

    success_count = 0
    fail_count = 0
    failures = []

    for i, ws in enumerate(affected_workspaces, 1):
        ws_id = ws["id"]
        ws_name = ws.get("displayName", "Unknown")

        # Build new license list: replace source with target in calling.webexCalling.licenses
        calling = ws.get("calling", {})
        webex_calling = calling.get("webexCalling", {})
        current_licenses = webex_calling.get("licenses", [])
        new_licenses = [target_id if lic == source_id else lic for lic in current_licenses]

        # PUT update — include displayName and calling with updated licenses
        # Must preserve locationId in webexCalling or API rejects it
        webex_calling_data = {"licenses": new_licenses}
        if webex_calling.get("locationId"):
            webex_calling_data["locationId"] = webex_calling["locationId"]
        elif ws.get("locationId"):
            webex_calling_data["locationId"] = ws["locationId"]

        update_data = {
            "displayName": ws.get("displayName", ""),
            "calling": {
                "type": "webexCalling",
                "webexCalling": webex_calling_data
            }
        }

        # Preserve optional fields if present
        if ws.get("capacity"):
            update_data["capacity"] = ws["capacity"]
        if ws.get("type"):
            update_data["type"] = ws["type"]
        if ws.get("notes"):
            update_data["notes"] = ws["notes"]

        result = api.call("PUT", f"workspaces/{ws_id}", data=update_data)

        if "error" in result:
            fail_count += 1
            failures.append((ws_name, result["error"]))
            status = "FAILED"
        else:
            success_count += 1
            status = "OK"

        # Progress counter
        print(f"  [{i}/{len(affected_workspaces)}] {ws_name:<40} {status}")

        # Rate limiting: brief pause every 10 requests
        if i % 10 == 0:
            time.sleep(1)

    # Step 7: Report
    print(f"\n{'=' * 60}")
    print(f"  RESULTS")
    print(f"{'=' * 60}")
    print(f"  Success: {success_count}")
    print(f"  Failed:  {fail_count}")
    print(f"{'=' * 60}")

    if success_count > 0:
        print(f"\n  Note: License usage counts may take a few minutes to update")
        print(f"  on Webex's backend. Re-run to see refreshed totals.")

    if failures:
        print(f"\n  Failed workspaces:")
        for name, error in failures[:20]:
            print(f"    - {name}: {error[:80]}")
        if len(failures) > 20:
            print(f"    ... and {len(failures) - 20} more (see API log for details)")
