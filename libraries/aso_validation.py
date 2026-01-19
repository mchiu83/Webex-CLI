# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

import re

def validate_excel_file(filepath):
    """Validate Excel file structure and required tabs"""
    import openpyxl
    import xlrd
    
    print(f"\nValidating Excel file: {filepath}")
    print(f"{'='*60}")
    
    required_tabs = ['Webex Users', 'Webex Side Cars', 'Webex Auto Attendant', 'Webex Hunt Groups']
    
    try:
        if filepath.endswith('.xlsx'):
            wb = openpyxl.load_workbook(filepath, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
        elif filepath.endswith('.xls'):
            wb = xlrd.open_workbook(filepath)
            sheet_names = wb.sheet_names()
        else:
            print(f"Status: FAILED - Unsupported file format")
            return False, None
        
        print(f"\nValidation 1: Checking required tabs...")
        all_required_present = True
        
        for tab in required_tabs:
            if tab in sheet_names:
                print(f"  [{tab}] - Status: PASS")
            else:
                print(f"  [{tab}] - Status: FAILED (Missing)")
                all_required_present = False
        
        if not all_required_present:
            print(f"\nOverall Status: FAILED - Missing required tabs")
            return False, None
        
        print(f"\nValidation 2: Checking for additional tabs...")
        additional_tabs = [tab for tab in sheet_names if tab not in required_tabs]
        
        if not additional_tabs:
            print(f"  Status: FAILED - No additional tabs found (at least one required)")
            return False, None
        else:
            print(f"  Status: PASS - Found {len(additional_tabs)} additional tab(s)")
            for tab in additional_tabs:
                print(f"    - {tab}")
        
        print(f"\n{'='*60}")
        print(f"Overall Validation Status: PASS")
        print(f"{'='*60}")
        
        return True, additional_tabs
        
    except Exception as e:
        print(f"\nStatus: FAILED - Error reading Excel file: {str(e)}")
        return False, None

def validate_location(api, filepath, read_excel_sheet):
    """Validation 3: Infer and validate location from Webex Users sheet"""
    print(f"\nValidation 3: Inferring and validating location...")
    
    users_data = read_excel_sheet(filepath, 'Webex Users')
    if not users_data:
        print(f"  Status: FAILED - Could not read 'Webex Users' sheet")
        return None
    
    headers = users_data[0]
    location_col_idx = None
    
    for idx, header in enumerate(headers):
        if header and 'Location Name' in str(header):
            location_col_idx = idx
            break
    
    if location_col_idx is None:
        print(f"  Status: FAILED - 'Location Name' column not found")
        return None
    
    if len(users_data) < 2:
        print(f"  Status: FAILED - No data rows found")
        return None
    
    inferred_location = None
    for row in users_data[1:]:
        if row[location_col_idx]:
            inferred_location = str(row[location_col_idx]).strip()
            break
    
    if not inferred_location:
        print(f"  Status: FAILED - No location name found")
        return None
    
    print(f"  Inferred Location: {inferred_location}")
    
    print(f"  Fetching telephony locations from Webex API...")
    locations_result = api.call("GET", "telephony/config/locations", params={"orgId": api.org_id})
    
    if "error" in locations_result:
        print(f"  Status: FAILED - Error fetching locations: {locations_result['error']}")
        return None
    
    locations = locations_result.get("locations", [])
    matched_location = None
    
    for loc in locations:
        if loc.get('name', '').lower() == inferred_location.lower():
            matched_location = loc
            break
    
    if not matched_location:
        print(f"  Status: FAILED - Location '{inferred_location}' not found")
        return None
    
    print(f"  Status: PASS - Location match found")
    print(f"  Location ID: {matched_location['id']}")
    print(f"  Location Name: {matched_location['name']}")
    
    calling_line_id = matched_location.get('callingLineId', {})
    phone_number = calling_line_id.get('phoneNumber')
    if phone_number:
        print(f"  Calling Line ID: {phone_number}")
    
    print(f"\n  Checking location outgoing permissions...")
    location_id = matched_location['id']
    perm_result = api.call("GET", f"telephony/config/locations/{location_id}/outgoingPermission", 
                          params={"orgId": api.org_id})
    
    if "error" in perm_result:
        print(f"  Status: FAILED - Error fetching outgoing permissions: {perm_result['error']}")
        return None
    
    permissions = perm_result.get('callingPermissions', [])
    
    print(f"\n  Location Outgoing Permissions:")
    print(f"  {'Call Type':<35} {'Action':<10} {'Transfer'}")
    print(f"  {'-'*55}")
    
    mismatches = []
    for perm in permissions:
        call_type = perm.get('callType', '')
        action = perm.get('action', '')
        transfer = perm.get('transferEnabled', False)
        print(f"  {call_type:<35} {action:<10} {transfer}")
        
        if call_type in ['CASUAL', 'URL_DIALING', 'UNKNOWN']:
            continue
        
        if call_type == 'INTERNAL_CALL':
            if action != 'ALLOW' or transfer != True:
                mismatches.append(call_type)
        else:
            if action != 'BLOCK' or transfer != False:
                mismatches.append(call_type)
    
    if not mismatches:
        print(f"\n  Status: PASS - Location permissions match expected configuration")
    else:
        print(f"\n  Status: WARNING - Location permissions do not match expected configuration")
        print(f"  Mismatched call types: {', '.join(mismatches)}")
        
        modify = input("\n  Modify location outgoing permissions to default? (y/n): ").strip().lower()
        if modify == 'y':
            print(f"  Updating location outgoing permissions...")
            
            update_data = {
                "callingPermissions": [
                    {"callType": "INTERNAL_CALL", "action": "ALLOW", "transferEnabled": True},
                    {"callType": "TOLL_FREE", "action": "BLOCK", "transferEnabled": False},
                    {"callType": "INTERNATIONAL", "action": "BLOCK", "transferEnabled": False},
                    {"callType": "OPERATOR_ASSISTED", "action": "BLOCK", "transferEnabled": False},
                    {"callType": "CHARGEABLE_DIRECTORY_ASSISTED", "action": "BLOCK", "transferEnabled": False},
                    {"callType": "SPECIAL_SERVICES_I", "action": "BLOCK", "transferEnabled": False},
                    {"callType": "SPECIAL_SERVICES_II", "action": "BLOCK", "transferEnabled": False},
                    {"callType": "PREMIUM_SERVICES_I", "action": "BLOCK", "transferEnabled": False},
                    {"callType": "PREMIUM_SERVICES_II", "action": "BLOCK", "transferEnabled": False},
                    {"callType": "NATIONAL", "action": "BLOCK", "transferEnabled": False}
                ]
            }
            
            update_result = api.call("PUT", f"telephony/config/locations/{location_id}/outgoingPermission",
                                    data=update_data, params={"orgId": api.org_id})
            
            if "error" in update_result:
                print(f"  Status: FAILED - Error updating permissions: {update_result['error']}")
                return None
            else:
                print(f"  Status: SUCCESS - Location outgoing permissions updated")
        else:
            print(f"  Proceeding without modifying location permissions...")
    
    return {
        'id': matched_location['id'],
        'name': matched_location['name'],
        'callingLineId': phone_number
    }

def validate_webex_users_data(filepath, read_excel_sheet):
    """Validation 4: Validate Webex Users sheet data"""
    print(f"\nValidation 4: Validating Webex Users data...")
    
    users_data = read_excel_sheet(filepath, 'Webex Users')
    if not users_data or len(users_data) < 2:
        print(f"  Status: FAILED - Could not read 'Webex Users' sheet or no data rows")
        return False
    
    headers = users_data[0]
    data_rows = users_data[1:]
    optional_cols = [0, 1, 3, 5, 6, 8, 13, 14, 15, 16, 17, 18]
    mac_addresses = []
    
    for row_idx, row in enumerate(data_rows, start=2):
        for col_idx in range(min(19, len(headers))):
            if col_idx not in optional_cols:
                if col_idx >= len(row) or not row[col_idx] or str(row[col_idx]).strip() == '':
                    col_name = str(headers[col_idx]).replace('\n', ' ').replace('\r', ' ') if col_idx < len(headers) else f"Column {chr(65 + col_idx)}"
                    print(f"  Status: FAILED - Row {row_idx}: Missing required value in '{col_name}'")
                    return False
        
        if len(row) <= 11 or not row[11] or str(row[11]).strip() == '':
            print(f"  Status: FAILED - Row {row_idx}: Missing MAC address")
            return False
        
        mac_raw = str(row[11]).strip()
        mac_clean = re.sub(r'[-:\s]', '', mac_raw).upper()
        if not re.match(r'^[0-9A-F]{12}$', mac_clean):
            print(f"  Status: FAILED - Row {row_idx}: Invalid MAC address format: '{mac_raw}'")
            return False
        
        mac_addresses.append(mac_clean)
        
        if len(row) <= 9 or not row[9]:
            print(f"  Status: FAILED - Row {row_idx}: Missing user type")
            return False
        
        col_j_value = str(row[9]).strip().lower()
        if col_j_value not in ['non-user', 'user']:
            print(f"  Status: FAILED - Row {row_idx}: User type must be 'non-user' or 'user', got '{row[9]}'")
            return False
        
        if len(row) > 4 and row[4] and str(row[4]).strip() != '':
            try:
                float(row[4])
            except ValueError:
                print(f"  Status: FAILED - Row {row_idx}: Extension must be numeric, got '{row[4]}'")
                return False
        
        if len(row) > 3 and row[3] and str(row[3]).strip() != '':
            col_d_value = str(row[3]).strip()
            if not col_d_value.isdigit() or len(col_d_value) != 10:
                print(f"  Status: FAILED - Row {row_idx}: Phone number must be 10 digits, got '{row[3]}'")
                return False
        
        if len(row) > 14 and row[14] and str(row[14]).strip() != '':
            try:
                col_o_value = float(row[14])
                if col_o_value > 15:
                    print(f"  Status: FAILED - Row {row_idx}: Rings must be <= 15, got '{row[14]}'")
                    return False
            except ValueError:
                print(f"  Status: FAILED - Row {row_idx}: Rings must be numeric, got '{row[14]}'")
                return False
        
        for col_idx, col_letter in [(15, 'P'), (17, 'R')]:
            if len(row) > col_idx and row[col_idx] and str(row[col_idx]).strip() != '':
                col_value = str(row[col_idx]).strip().lower()
                if col_value not in ['yes', 'no']:
                    print(f"  Status: FAILED - Row {row_idx}: Column {col_letter} must be 'yes', 'no', or empty")
                    return False
        
        for col_idx, col_letter in [(13, 'N'), (16, 'Q')]:
            if len(row) > col_idx and row[col_idx] and str(row[col_idx]).strip() != '':
                try:
                    float(row[col_idx])
                except ValueError:
                    print(f"  Status: FAILED - Row {row_idx}: Column {col_letter} must be numeric")
                    return False
    
    if len(mac_addresses) != len(set(mac_addresses)):
        duplicates = [mac for mac in set(mac_addresses) if mac_addresses.count(mac) > 1]
        print(f"  Status: FAILED - Duplicate MAC addresses found: {', '.join(duplicates)}")
        return False
    
    print(f"  Status: PASS - All {len(data_rows)} rows validated successfully")
    return True

def validate_available_numbers(api, location_data, filepath, read_excel_sheet):
    """Validation 5: Validate phone numbers against available location numbers"""
    print(f"\nValidation 5: Validating phone number availability...")
    
    print(f"  Fetching available numbers for location...")
    location_id = location_data['id']
    numbers_result = api.call("GET", f"telephony/config/locations/{location_id}/availableNumbers", 
                             params={"orgId": api.org_id})
    
    if "error" in numbers_result:
        print(f"  Status: FAILED - Error fetching available numbers: {numbers_result['error']}")
        return False
    
    phone_numbers = numbers_result.get("phoneNumbers", [])
    available_location_numbers = []
    
    for num in phone_numbers:
        if (not num.get('owner') and 
            not num.get('isMainNumber', False) and 
            num.get('state') == 'ACTIVE'):
            available_location_numbers.append(num['phoneNumber'])
    
    print(f"  Found {len(available_location_numbers)} available numbers")
    
    users_data = read_excel_sheet(filepath, 'Webex Users')
    if not users_data or len(users_data) < 2:
        print(f"  Status: FAILED - Could not read 'Webex Users' sheet")
        return False
    
    headers = users_data[0]
    data_rows = users_data[1:]
    
    for row_idx, row in enumerate(data_rows, start=2):
        if len(row) > 3 and row[3] and str(row[3]).strip() != '':
            phone_10digit = str(row[3]).strip()
            phone_e164 = f"+1{phone_10digit}"
            
            if phone_e164 not in available_location_numbers:
                print(f"  Status: FAILED - Row {row_idx}: Phone number '{phone_10digit}' is not available")
                print(f"  The number {phone_e164} is not found in available PSTN numbers for this location.")
                return False
    
    print(f"  Status: PASS - All phone numbers are available")
    return True
