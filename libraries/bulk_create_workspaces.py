# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

import csv
import os
import re
from libraries.add_device import PHONE_MODELS, COLLAB_MODELS

def validate_csv_structure(filepath):
    """Validate CSV file structure and format"""
    if not os.path.exists(filepath):
        return False, f"File not found: {filepath}"
    
    try:
        with open(filepath, 'r') as f:
            # Check if file is empty
            content = f.read()
            if not content.strip():
                return False, "CSV file is empty"
            
            # Reset file pointer
            f.seek(0)
            
            # Try to parse as CSV
            reader = csv.DictReader(f)
            
            # Validate headers
            expected_headers = ['id', 'location', 'displayName', 'supportedDevices', 'type', 
                              'capacity', 'calling', 'extension', 'phoneNumber', 'phoneModel', 'macaddress']
            
            if not reader.fieldnames:
                return False, "CSV file has no headers"
            
            missing_headers = set(expected_headers) - set(reader.fieldnames)
            if missing_headers:
                return False, f"Missing required columns: {', '.join(missing_headers)}"
            
            # Validate each row has correct number of fields
            f.seek(0)
            next(f)  # Skip header
            for i, line in enumerate(f, start=2):
                field_count = len(line.split(','))
                if field_count != len(expected_headers):
                    return False, f"Row {i}: Invalid number of fields (expected {len(expected_headers)}, got {field_count})"
        
        return True, "CSV structure is valid"
    
    except Exception as e:
        return False, f"Error reading CSV: {str(e)}"

def validate_workspace_data(row, row_num, available_locations):
    """Validate individual workspace row data"""
    errors = []
    
    # Validate displayName (mandatory)
    if not row.get('displayName', '').strip():
        errors.append(f"Row {row_num}: displayName is mandatory")
    
    # Validate supportedDevices
    supported_devices = row.get('supportedDevices', '').strip().lower()
    if not supported_devices:
        supported_devices = 'collaborationdevices'
    elif supported_devices not in ['phones', 'collaborationdevices']:
        errors.append(f"Row {row_num}: supportedDevices must be 'phones' or 'collaborationDevices', got '{row.get('supportedDevices')}'")
    
    # Validate calling
    calling = row.get('calling', '').strip().lower()
    if not calling:
        calling = 'none'
    elif calling not in ['none', 'webexcalling']:
        errors.append(f"Row {row_num}: calling must be 'none' or 'webexCalling', got '{row.get('calling')}'")
    
    # Validate extension if webexCalling
    if calling == 'webexcalling':
        extension = row.get('extension', '').strip()
        if not extension:
            errors.append(f"Row {row_num}: extension is mandatory when calling is 'webexCalling'")
        elif not extension.isdigit() or len(extension) < 4:
            errors.append(f"Row {row_num}: extension must be at least 4 digits and contain only numbers, got '{extension}'")
        
        # Validate phoneNumber (optional)
        phone_number = row.get('phoneNumber', '').strip()
        if phone_number:
            if not phone_number.isdigit() or len(phone_number) != 10:
                errors.append(f"Row {row_num}: phoneNumber must be exactly 10 digits, got '{phone_number}'")
        
        # Validate phoneModel
        phone_model = row.get('phoneModel', '').strip()
        if phone_model:
            valid_models = PHONE_MODELS if supported_devices == 'phones' else COLLAB_MODELS
            if phone_model not in valid_models:
                errors.append(f"Row {row_num}: phoneModel '{phone_model}' is not valid for {supported_devices}")
        
        # Validate macaddress
        mac_address = row.get('macaddress', '').strip()
        if mac_address:
            # Remove any separators and validate
            mac_clean = ''.join(c for c in mac_address if c.isalnum())
            if len(mac_clean) != 12:
                errors.append(f"Row {row_num}: macaddress must be 12 alphanumeric characters without separators, got '{mac_address}'")
            elif not re.match(r'^[0-9A-Fa-f]{12}$', mac_clean):
                errors.append(f"Row {row_num}: macaddress contains invalid characters, got '{mac_address}'")
    else:
        # If calling is none, phoneModel and macaddress should be empty
        if row.get('phoneModel', '').strip():
            errors.append(f"Row {row_num}: phoneModel should only be populated when calling is 'webexCalling'")
    
    # Validate location if provided
    location = row.get('location', '').strip()
    if location and location not in [loc['name'] for loc in available_locations]:
        errors.append(f"Row {row_num}: location '{location}' not found in available locations")
    
    return errors, supported_devices, calling

def parse_workspaces_csv(api):
    """Parse and validate workspaces.csv file"""
    filepath = "bulk/workspaces.csv"
    
    # Validate CSV structure
    is_valid, message = validate_csv_structure(filepath)
    if not is_valid:
        print(f"CSV Validation Error: {message}")
        return None
    
    print("CSV structure validated successfully.")
    
    # Get available locations
    locations_result = api.call("GET", "locations", params={"orgId": api.org_id})
    if "error" in locations_result:
        print(f"Error fetching locations: {locations_result['error']}")
        return None
    
    available_locations = locations_result.get("items", [])
    
    # Parse and validate each row
    workspaces = []
    all_errors = []
    
    with open(filepath, 'r') as f:
        reader = csv.DictReader(f)
        for i, row in enumerate(reader, start=2):
            errors, supported_devices, calling = validate_workspace_data(row, i, available_locations)
            
            if errors:
                all_errors.extend(errors)
            else:
                # Store validated workspace data
                workspace_data = {
                    'row_num': i,
                    'displayName': row['displayName'].strip(),
                    'supportedDevices': supported_devices,
                    'type': row.get('type', '').strip() or 'notSet',
                    'capacity': row.get('capacity', '').strip(),
                    'calling': calling,
                    'location': row.get('location', '').strip(),
                    'extension': row.get('extension', '').strip(),
                    'phoneNumber': row.get('phoneNumber', '').strip(),
                    'phoneModel': row.get('phoneModel', '').strip(),
                    'macaddress': row.get('macaddress', '').strip()
                }
                workspaces.append(workspace_data)
    
    if all_errors:
        print("\nValidation Errors Found:")
        for error in all_errors:
            print(f"  - {error}")
        print("\nPlease fix the errors in the CSV file before proceeding.")
        return None
    
    return workspaces, available_locations

def display_workspace_summary(workspaces):
    """Display summary table of workspaces to be created"""
    print(f"\n{'='*160}")
    print(f"{'Row':<5} {'Location':<25} {'Display Name':<25} {'Devices':<25} {'Calling':<15} {'Extension':<10} {'Phone':<12} {'Model':<20} {'MAC':<20}")
    print(f"{'='*160}")
    
    for ws in workspaces:
        calling_display = ws['calling'] if ws['calling'] != 'none' else 'None'
        extension_display = ws['extension'] if ws['calling'] == 'webexcalling' else '-'
        phone_display = ws['phoneNumber'] if ws['phoneNumber'] else '-'
        model_display = ws['phoneModel'] if ws['phoneModel'] else '-'
        mac_display = ws['macaddress'] if ws['macaddress'] else '-'

        print(f"{ws['row_num'] - 1:<5} {ws['location']:<25} {ws['displayName']:<25} {ws['supportedDevices']:<25} {calling_display:<15} {extension_display:<10} {phone_display:<12} {model_display:<20} {mac_display:<20}")
    
    print(f"{'='*160}")

def execute_bulk_create(api, workspaces, available_locations):
    """Execute bulk workspace creation"""
    print(f"\nStarting bulk creation of {len(workspaces)} workspace(s)...")
    
    results = []
    
    for ws in workspaces:
        print(f"\nCreating workspace: {ws['displayName']} (Row {ws['row_num']})")
        
        # Prepare workspace data
        data = {
            "displayName": ws['displayName'],
            "orgId": api.org_id,
            "type": ws['type'],
            "supportedDevices": ws['supportedDevices']
        }
        
        if ws['capacity']:
            data['capacity'] = int(ws['capacity'])
        
        # Handle location and calling
        if ws['calling'] == 'webexcalling':
            # Get location ID
            if ws['location']:
                location = next((loc for loc in available_locations if loc['name'] == ws['location']), None)
                if location:
                    location_id = location['id']
                else:
                    print(f"  Error: Location '{ws['location']}' not found")
                    results.append({'row': ws['row_num'], 'name': ws['displayName'], 'status': 'failed', 'error': 'Location not found'})
                    continue
            else:
                # Ask user for location
                print("\n  Available Locations:")
                for i, loc in enumerate(available_locations, 1):
                    print(f"  {i}. {loc.get('name', 'N/A')}")
                
                loc_choice = input("  Select location number: ").strip()
                try:
                    location_id = available_locations[int(loc_choice) - 1]["id"]
                except (ValueError, IndexError):
                    print("  Invalid selection. Skipping workspace.")
                    results.append({'row': ws['row_num'], 'name': ws['displayName'], 'status': 'failed', 'error': 'Invalid location selection'})
                    continue
            
            data['locationId'] = location_id
            data['calling'] = {
                "type": "webexCalling",
                "webexCalling": {
                    "extension": ws['extension'],
                    "locationId": location_id
                }
            }
            
            if ws['phoneNumber']:
                data['calling']['webexCalling']['phoneNumber'] = ws['phoneNumber']
        
        # Create workspace
        result = api.call("POST", "workspaces", data=data)
        
        if "error" in result:
            print(f"  Error creating workspace: {result['error']}")
            results.append({'row': ws['row_num'], 'name': ws['displayName'], 'status': 'failed', 'error': result['error']})
            continue
        
        workspace_id = result.get("id")
        print(f"  Workspace created successfully! ID: {workspace_id}")
        
        # Create device if phoneModel is specified
        if ws['phoneModel'] and ws['calling'] == 'webexcalling':
            print(f"  Creating device: {ws['phoneModel']}")
            
            if ws['macaddress']:
                # Create with MAC address
                mac_clean = ''.join(c for c in ws['macaddress'].upper() if c.isalnum())
                mac_formatted = ':'.join(mac_clean[i:i+2] for i in range(0, 12, 2))
                
                device_data = {
                    "mac": mac_formatted,
                    "model": ws['phoneModel'],
                    "workspaceId": workspace_id
                }
                device_result = api.call("POST", "devices", data=device_data, params={"orgId": api.org_id})
            else:
                # Create with activation code
                device_data = {
                    "workspaceId": workspace_id,
                    "model": ws['phoneModel']
                }
                device_result = api.call("POST", "devices/activationCode", data=device_data, params={"orgId": api.org_id})
            
            if "error" in device_result:
                print(f"  Warning: Device creation failed: {device_result['error']}")
                results.append({'row': ws['row_num'], 'name': ws['displayName'], 'status': 'partial', 'workspace_id': workspace_id, 'error': f"Device creation failed: {device_result['error']}"})
            else:
                activation_code = device_result.get('code', 'N/A') if not ws['macaddress'] else 'MAC'
                print(f"  Device created successfully! Activation: {activation_code}")
                results.append({'row': ws['row_num'], 'name': ws['displayName'], 'status': 'success', 'workspace_id': workspace_id})
        else:
            results.append({'row': ws['row_num'], 'name': ws['displayName'], 'status': 'success', 'workspace_id': workspace_id})
    
    # Display results summary
    print(f"\n{'='*100}")
    print("Bulk Creation Results:")
    print(f"{'='*100}")
    print(f"{'Row':<5} {'Display Name':<30} {'Status':<10} {'Workspace ID / Error':<50}")
    print(f"{'='*100}")
    
    for r in results:
        status_display = r['status'].upper()
        detail = r.get('workspace_id', r.get('error', ''))
        print(f"{r['row']:<5} {r['name']:<30} {status_display:<10} {detail:<50}")
    
    print(f"{'='*100}")
    
    success_count = sum(1 for r in results if r['status'] == 'success')
    partial_count = sum(1 for r in results if r['status'] == 'partial')
    failed_count = sum(1 for r in results if r['status'] == 'failed')
    
    print(f"\nTotal: {len(results)} | Success: {success_count} | Partial: {partial_count} | Failed: {failed_count}")

def bulk_create_workspaces(api):
    """Main function for bulk workspace creation"""
    print("\n--- Bulk Create Workspaces ---")
    
    # Check if bulk folder exists
    if not os.path.exists("bulk"):
        print("Error: 'bulk' folder not found. Creating it now...")
        os.makedirs("bulk")
        print("Please place your 'workspaces.csv' file in the 'bulk' folder and try again.")
        return
    
    # Parse and validate CSV
    result = parse_workspaces_csv(api)
    if result is None:
        return
    
    workspaces, available_locations = result
    
    if not workspaces:
        print("No valid workspaces found in CSV file.")
        return
    
    # Display summary
    print(f"\nBulk admin will create {len(workspaces)} workspace(s).")
    
    while True:
        choice = input("\nOptions: (p)roceed, (d)etails, (c)ancel: ").strip().lower()
        
        if choice == 'p':
            execute_bulk_create(api, workspaces, available_locations)
            break
        elif choice == 'd':
            display_workspace_summary(workspaces)
        elif choice == 'c':
            print("Bulk creation cancelled.")
            break
        else:
            print("Invalid choice. Please enter 'p', 'd', or 'c'.")
