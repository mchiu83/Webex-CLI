# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

import os
import glob
import openpyxl
import xlrd

def find_aso_import_file():
    """Find Excel file with prefix 'aso_import' in bulk directory"""
    bulk_dir = 'bulk'
    
    if not os.path.exists(bulk_dir):
        return None
    
    patterns = [os.path.join(bulk_dir, 'aso_import*.xlsx'), 
                os.path.join(bulk_dir, 'aso_import*.xls')]
    
    for pattern in patterns:
        files = glob.glob(pattern)
        if files:
            return files[0]
    
    return None


def _is_valid_import_file(filepath: str) -> tuple:
    """
    Check if an Excel file is a valid bulk import file.
    Criteria:
      - Has tabs: Webex Users, Webex Side Cars, Webex Auto Attendant, Webex Hunt Groups
      - First sheet A1 contains store location info (newline-separated address)
    Returns (is_valid: bool, address_lines: list[str])
    """
    required_tabs = ['Webex Users', 'Webex Side Cars', 'Webex Auto Attendant', 'Webex Hunt Groups']
    try:
        import warnings
        warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
        if filepath.endswith('.xlsx'):
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            sheet_names = wb.sheetnames
            if not all(tab in sheet_names for tab in required_tabs):
                wb.close()
                return False, []
            first_ws = wb[sheet_names[0]]
            a1 = None
            for row in first_ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=1, values_only=True):
                a1 = row[0]
            wb.close()
        elif filepath.endswith('.xls'):
            wb = xlrd.open_workbook(filepath)
            sheet_names = wb.sheet_names()
            if not all(tab in sheet_names for tab in required_tabs):
                return False, []
            first_ws = wb.sheet_by_index(0)
            a1 = first_ws.cell_value(0, 0) if first_ws.nrows > 0 else None
        else:
            return False, []

        if not a1:
            return False, []
        lines = [l.strip() for l in str(a1).split('\n') if l.strip()]
        if not lines:
            return False, []
        return True, lines

    except Exception:
        return False, []


def select_import_file() -> str | None:
    """
    Scan the /bulk folder for valid import files, present a formatted table,
    and return the selected filepath. Returns None if none found or cancelled.
    """
    bulk_dir = 'bulk'
    if not os.path.exists(bulk_dir):
        return None

    candidates = []
    for pattern in [os.path.join(bulk_dir, '*.xlsx'), os.path.join(bulk_dir, '*.xls')]:
        for f in sorted(glob.glob(pattern)):
            if os.path.basename(f).startswith('~$'):
                continue
            is_valid, address_lines = _is_valid_import_file(f)
            if is_valid:
                candidates.append((f, address_lines))

    if not candidates:
        print("  No valid import files found in /bulk folder.")
        print("  A valid file must contain tabs: Webex Users, Webex Side Cars, Webex Auto Attendant, Webex Hunt Groups")
        return None

    if len(candidates) == 1:
        filepath, address_lines = candidates[0]
        print(f"  Found 1 valid import file: {os.path.basename(filepath)}")
        if address_lines:
            print(f"  Store: {address_lines[0]}")
        return filepath

    # Multiple files - present formatted table
    col_num_w  = 1
    col_file_w = max(len(os.path.basename(f)) for f, _ in candidates)
    col_addr_w = max((max(len(ln) for ln in addr) if addr else 0) for _, addr in candidates)
    col_addr_w = max(col_addr_w, len("Store Information"))
    col_file_w = max(col_file_w, len("File Name"))

    sep = f"  +{chr(45)*(col_num_w+2)}+{chr(45)*(col_file_w+2)}+{chr(45)*(col_addr_w+2)}+"

    print(f"\n{sep}")
    print("  | " + "#".ljust(col_num_w) + " | " + "File Name".ljust(col_file_w) + " | " + "Store Information".ljust(col_addr_w) + " |")
    print(f"{sep}")

    for i, (filepath, address_lines) in enumerate(candidates, 1):
        filename = os.path.basename(filepath)
        first_line = address_lines[0] if address_lines else ""
        print("  | " + str(i).ljust(col_num_w) + " | " + filename.ljust(col_file_w) + " | " + first_line.ljust(col_addr_w) + " |")
        for line in address_lines[1:]:
            print("  | " + "".ljust(col_num_w) + " | " + "".ljust(col_file_w) + " | " + line.ljust(col_addr_w) + " |")
        if i < len(candidates):
            print(f"{sep}")

    print(f"{sep}")

    while True:
        choice = input(f"\n  Select file (1-{len(candidates)}): ").strip()
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(candidates):
                return candidates[idx][0]
        except ValueError:
            pass
        print(f"  Invalid selection. Please enter a number between 1 and {len(candidates)}.")

def read_excel_sheet(filepath, sheet_name):
    """Read data from specific Excel sheet"""
    try:
        if filepath.endswith('.xlsx'):
            import warnings
            warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            ws = wb[sheet_name]
            data = []
            for row in ws.iter_rows(values_only=True):
                data.append(row)
            wb.close()
            return data
        elif filepath.endswith('.xls'):
            wb = xlrd.open_workbook(filepath, formatting_info=False)
            ws = wb.sheet_by_name(sheet_name)
            data = []
            for row_idx in range(ws.nrows):
                data.append(ws.row_values(row_idx))
            return data
    except Exception as e:
        print(f"Error reading sheet '{sheet_name}': {str(e)}")
        return None

def check_mac_conflicts(api, preview_items):
    """
    Validate MAC addresses against Control Hub using the validateMacs endpoint.
    Returns True if all MACs are clean, False if any conflicts are found.
    """
    import re

    # Collect MACs from workspace rows only (users don't get devices)
    mac_map = {}  # formatted_mac -> row number
    for item in preview_items:
        if item['type'] == 'Workspace' and item['mac']:
            mac_raw = item['mac']
            mac_clean = re.sub(r'[-:\s]', '', mac_raw).upper()
            if re.match(r'^[0-9A-F]{12}$', mac_clean):
                mac_formatted = ':'.join(mac_clean[i:i+2] for i in range(0, 12, 2))
                mac_map[mac_formatted] = item['row']

    if not mac_map:
        return True

    print(f"\nChecking {len(mac_map)} MAC address(es) against Control Hub...")

    result = api.call(
        "POST",
        "telephony/config/devices/actions/validateMacs/invoke",
        data={"macs": list(mac_map.keys())}
    )

    if "error" in result:
        print(f"  Warning: MAC validation API call failed ({result['error']}) - skipping conflict check")
        return True

    mac_info = result.get("macs", [])
    conflicts = []

    for entry in mac_info:
        mac = entry.get("mac", "")
        status = entry.get("state", "")
        message = entry.get("message", "")
        row = mac_map.get(mac, "?")

        if status == "available":
            print(f"  Row {row}: {mac} - OK")
        else:
            # Any non-available state (e.g. "duplicate", "invalid") is a conflict
            print(f"  Row {row}: {mac} - CONFLICT ({message or status})")
            conflicts.append((row, mac, message or status))

    if conflicts:
        print(f"\n{'='*97}")
        print(f"MAC Conflict Check FAILED - {len(conflicts)} conflict(s) found. Resolve before importing.")
        print(f"{'='*97}")
        return False

    print(f"  MAC conflict check passed - all addresses are available.")
    return True


def process_bulk_import(api, location_data, filepath):
    """Process bulk import of workspaces from Excel file"""
    from libraries.workspace_config import (
        create_workspace_from_row,
        configure_call_forwarding,
        configure_outgoing_permission,
        configure_side_car_speed_dials
    )
    
    users_data = read_excel_sheet(filepath, 'Webex Users')
    if not users_data or len(users_data) < 2:
        print("Error: Could not read data")
        return
    
    headers = users_data[0]
    data_rows = users_data[1:]
    
    print(f"\n{'='*80}")
    print("Import Preview")
    print(f"{'='*80}")
    
    users_count = 0
    workspaces_count = 0
    preview_items = []
    
    for row_idx, row in enumerate(data_rows, start=2):
        user_type = str(row[9]).strip().lower() if len(row) > 9 else ""
        display_name = str(row[12]).strip() if len(row) > 12 else "Unknown"
        extension = str(row[4]).strip() if len(row) > 4 else ""
        phone_number = str(row[3]).strip() if len(row) > 3 and row[3] else ""
        device_model = str(row[10]).strip() if len(row) > 10 else ""
        mac_address = str(row[11]).strip() if len(row) > 11 and row[11] else ""
        
        if user_type == 'user':
            users_count += 1
            preview_items.append({
                'row': row_idx,
                'type': 'User',
                'name': display_name,
                'ext': extension,
                'phone': phone_number,
                'device': device_model,
                'mac': mac_address
            })
        else:
            workspaces_count += 1
            preview_items.append({
                'row': row_idx,
                'type': 'Workspace',
                'name': display_name,
                'ext': extension,
                'phone': phone_number,
                'device': device_model,
                'mac': mac_address
            })
    
    print(f"{'Row':<5} {'Type':<10} {'Name':<25} {'Ext':<8} {'Phone':<12} {'Device':<20} {'MAC Address':<17}")
    print(f"{'-'*97}")
    for item in preview_items:
        print(f"{item['row']:<5} {item['type']:<10} {item['name']:<25} {item['ext']:<8} {item['phone']:<12} {item['device']:<20} {item['mac']:<17}")
    
    print(f"\n{'='*97}")
    print(f"Total: {len(preview_items)} items ({users_count} users, {workspaces_count} workspaces)")
    print(f"Note: Users will be skipped (not yet implemented)")
    print(f"{'='*97}")

    if not check_mac_conflicts(api, preview_items):
        return

    confirm = input("\nProceed with import? (Y/n): ").strip().lower()
    if confirm not in ['', 'y', 'yes']:
        print("Import cancelled.")
        return
    
    print(f"\n{'='*60}")
    print("Starting Bulk Import Process")
    print(f"{'='*60}")
    
    results = {'users': 0, 'workspaces_created': 0, 'workspaces_failed': 0, 'errors': []}
    workspace_map = {}
    
    for row_idx, row in enumerate(data_rows, start=2):
        user_type = str(row[9]).strip().lower() if len(row) > 9 else ""
        display_name = str(row[12]).strip() if len(row) > 12 else "Unknown"
        
        if user_type == 'user':
            results['users'] += 1
            print(f"Row {row_idx}: Skipping user '{display_name}' (user provisioning not yet implemented)")
            continue
        
        print(f"\nRow {row_idx}: Creating workspace '{display_name}'...")
        workspace_id, error = create_workspace_from_row(api, location_data, row, headers)
        
        if workspace_id:
            results['workspaces_created'] += 1
            workspace_map[row_idx] = workspace_id
            if error:
                print(f"  Warning: {error}")
                results['errors'].append(f"Row {row_idx}: {error}")
            else:
                print(f"  Success: Workspace created (ID: {workspace_id})")
        else:
            results['workspaces_failed'] += 1
            print(f"  Failed: {error}")
            results['errors'].append(f"Row {row_idx}: {error}")
    
    if workspace_map:
        print(f"\n{'='*60}")
        print("Configuring Call Forwarding & Business Continuity")
        print(f"{'='*60}")
        
        for row_idx, workspace_id in workspace_map.items():
            row = data_rows[row_idx - 2]
            display_name = str(row[12]).strip() if len(row) > 12 else "Unknown"
            
            print(f"\nRow {row_idx}: Configuring '{display_name}'...")
            error = configure_call_forwarding(api, workspace_id, row)
            
            if error:
                print(f"  Warning: {error}")
                results['errors'].append(f"Row {row_idx}: Call forwarding failed - {error}")
            else:
                print(f"  Success: Call forwarding configured")
    
    if workspace_map:
        print(f"\n{'='*60}")
        print("Configuring Outgoing Calling Permissions")
        print(f"{'='*60}")
        
        for row_idx, workspace_id in workspace_map.items():
            row = data_rows[row_idx - 2]
            display_name = str(row[12]).strip() if len(row) > 12 else "Unknown"
            
            print(f"\nRow {row_idx}: Checking '{display_name}'...")
            error, was_configured = configure_outgoing_permission(api, workspace_id, row)
            
            if was_configured:
                if error:
                    print(f"  Warning: {error}")
                    results['errors'].append(f"Row {row_idx}: Outgoing permission failed - {error}")
                else:
                    print(f"  Success: Custom outgoing permissions configured")
            else:
                print(f"  Skipped: No custom permissions required")
    
    if workspace_map:
        print(f"\n{'='*60}")
        proceed = input("\nProceed with side car speed dial configuration? (Y/n): ").strip().lower()
        if proceed in ['', 'y', 'yes']:
            configure_side_car_speed_dials(api, workspace_map, data_rows, filepath, read_excel_sheet)
        else:
            print("\nSide car configuration skipped.")
    
    if workspace_map:
        from libraries.configure_hunt_groups import configure_hunt_groups
        configure_hunt_groups(api, location_data, workspace_map, data_rows, filepath)

    from libraries.configure_call_park_group import configure_call_park_group
    configure_call_park_group(api, location_data, filepath, read_excel_sheet)

    from libraries.configure_auto_attendant import configure_auto_attendants
    configure_auto_attendants(api, location_data, filepath, read_excel_sheet)

    print(f"\n{'='*60}")
    print("Bulk Import Summary")
    print(f"{'='*60}")
    print(f"Users skipped: {results['users']}")
    print(f"Workspaces created: {results['workspaces_created']}")
    print(f"Workspaces failed: {results['workspaces_failed']}")
    
    if results['errors']:
        print(f"\nErrors/Warnings:")
        for error in results['errors']:
            print(f"  - {error}")
    
    if results['users'] > 0:
        print(f"\nNote: User provisioning will be implemented in a future update.")
    
    print(f"{'='*60}")

def aso_bulk_import_tool(api):
    """Main function for ASO Bulk Import Tool"""
    from libraries.aso_validation import (
        validate_excel_file,
        validate_location,
        validate_webex_users_data,
        validate_available_numbers,
        validate_translation_pattern,
        validate_call_park_extensions
    )
    from libraries.schedule_manager import validate_and_create_schedules
    
    print("\n--- ASO Bulk Import Tool ---")
    
    print("\nSearching for valid import files in bulk folder...")
    filepath = select_import_file()
    
    if not filepath:
        print("Status: FAILED - No valid import file selected.")
        print("\nPlease ensure your Excel file:")
        print("  1. Is in Excel format (.xlsx or .xls)")
        print("  2. Is located in the 'bulk' folder")
        print("  3. Contains tabs: Webex Users, Webex Side Cars, Webex Auto Attendant, Webex Hunt Groups")
        return
    
    print(f"Status: PASS - Using file: {filepath}")
    
    is_valid, additional_tabs = validate_excel_file(filepath)
    
    if not is_valid:
        print("\nValidation failed. Please fix the issues and try again.")
        return
    
    if additional_tabs:
        print(f"\nAdditional tabs detected and cached:")
        for i, tab in enumerate(additional_tabs, 1):
            print(f"  {i}. {tab}")
    
    location = validate_location(api, filepath, read_excel_sheet)
    
    if not location:
        print("\nValidation failed. Returning to previous menu.")
        return
    
    if not validate_webex_users_data(filepath, read_excel_sheet):
        print("\nValidation failed. Returning to previous menu.")
        return
    
    if not validate_available_numbers(api, location, filepath, read_excel_sheet):
        print("\nValidation failed. Returning to previous menu.")
        return
    
    translation_pattern = validate_translation_pattern(api, location, filepath, read_excel_sheet, additional_tabs)
    
    call_park_extensions = validate_call_park_extensions(api, location, filepath, read_excel_sheet, additional_tabs)
    
    schedule_ids = validate_and_create_schedules(api, location['id'], filepath)
    
    print("\nValidation complete. Ready for next steps.")
    
    process_bulk_import(api, location, filepath)
