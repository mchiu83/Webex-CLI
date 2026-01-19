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
        
        if user_type == 'user':
            users_count += 1
            preview_items.append({
                'row': row_idx,
                'type': 'User',
                'name': display_name,
                'ext': extension,
                'phone': phone_number,
                'device': device_model
            })
        else:
            workspaces_count += 1
            preview_items.append({
                'row': row_idx,
                'type': 'Workspace',
                'name': display_name,
                'ext': extension,
                'phone': phone_number,
                'device': device_model
            })
    
    print(f"{'Row':<5} {'Type':<10} {'Name':<25} {'Ext':<8} {'Phone':<12} {'Device':<20}")
    print(f"{'-'*80}")
    for item in preview_items:
        print(f"{item['row']:<5} {item['type']:<10} {item['name']:<25} {item['ext']:<8} {item['phone']:<12} {item['device']:<20}")
    
    print(f"\n{'='*80}")
    print(f"Total: {len(preview_items)} items ({users_count} users, {workspaces_count} workspaces)")
    print(f"Note: Users will be skipped (not yet implemented)")
    print(f"{'='*80}")
    
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
    
    if not os.path.exists('bulk'):
        print("Status: FAILED - 'bulk' folder not found")
        print("Creating 'bulk' folder...")
        os.makedirs('bulk')
        print("Please place your 'aso_import' Excel file in the 'bulk' folder and try again.")
        return
    
    print("\nSearching for 'aso_import' Excel file in bulk folder...")
    filepath = find_aso_import_file()
    
    if not filepath:
        print("Status: FAILED - No file found with prefix 'aso_import' (.xlsx or .xls)")
        print("\nPlease ensure your Excel file:")
        print("  1. Has a filename starting with 'aso_import'")
        print("  2. Is in Excel format (.xlsx or .xls)")
        print("  3. Is located in the 'bulk' folder")
        return
    
    print(f"Status: PASS - Found file: {filepath}")
    
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
