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

def validate_translation_pattern(api, location_data, filepath, read_excel_sheet, additional_tabs):
    """Validation 6: Validate organization translation pattern"""
    import re
    
    print(f"\nValidation 6: Validating organization translation pattern...")
    
    # Find location-specific tab
    location_name = location_data['name']
    location_tab = None
    
    for tab in additional_tabs:
        if tab.lower() == location_name.lower():
            location_tab = tab
            break
    
    if not location_tab:
        print(f"  Status: FAILED - Location tab '{location_name}' not found in additional tabs")
        print(f"  Available tabs: {', '.join(additional_tabs)}")
        input("  Press Enter to acknowledge and continue...")
        return {}
    
    print(f"  Found location tab: {location_tab}")
    
    # Read location tab
    location_data_sheet = read_excel_sheet(filepath, location_tab)
    if not location_data_sheet or len(location_data_sheet) < 65:
        print(f"  Status: FAILED - Could not read location tab or insufficient rows")
        input("  Press Enter to acknowledge and continue...")
        return {}
    
    # Extract translation pattern data from B62, B63, B64 (column index 1, rows 61-63)
    translation_name = str(location_data_sheet[61][1]).strip() if len(location_data_sheet[61]) > 1 and location_data_sheet[61][1] else ""
    matching_pattern = str(location_data_sheet[62][1]).strip() if len(location_data_sheet[62]) > 1 and location_data_sheet[62][1] else ""
    replacement_pattern = str(location_data_sheet[63][1]).strip() if len(location_data_sheet[63]) > 1 and location_data_sheet[63][1] else ""
    
    if not translation_name or not matching_pattern or not replacement_pattern:
        print(f"  Status: WARNING - Translation pattern data missing in location tab")
        print(f"    B62 (Name): {'[MISSING]' if not translation_name else translation_name}")
        print(f"    B63 (Matching Pattern): {'[MISSING]' if not matching_pattern else matching_pattern}")
        print(f"    B64 (Replacement Pattern): {'[MISSING]' if not replacement_pattern else replacement_pattern}")
        print(f"  Please fix the translation pattern data in Excel tab '{location_tab}'")
        input("  Press Enter to acknowledge and continue...")
        return {}
    
    # Display translation pattern table
    print(f"\n  Translation Pattern Information:")
    print(f"  {'Field':<20} {'Value':<40}")
    print(f"  {'-'*60}")
    print(f"  {'Name':<20} {translation_name:<40}")
    print(f"  {'Matching Pattern':<20} {matching_pattern:<40}")
    print(f"  {'Replacement Pattern':<20} {replacement_pattern:<40}")
    
    # Query existing translation patterns
    print(f"\n  Checking for existing translation pattern...")
    result = api.call("GET", "telephony/config/callRouting/translationPatterns",
                     params={"orgId": api.org_id, "matchingPattern": matching_pattern})
    
    if "error" in result:
        print(f"  Status: FAILED - Error fetching translation patterns: {result['error']}")
        input("  Press Enter to continue...")
        return {}
    
    translation_patterns = result.get('translationPatterns', [])
    
    if translation_patterns:
        # Check if matching pattern matches
        found_pattern = translation_patterns[0]
        if found_pattern.get('matchingPattern') == matching_pattern:
            print(f"  Status: PASS - Translation pattern exists")
            print(f"  Pattern ID: {found_pattern.get('id')}")
            print(f"  Pattern Name: {found_pattern.get('name')}")
            input("\n  Press Enter to proceed to next step...")
            return {'id': found_pattern.get('id'), 'name': found_pattern.get('name')}
        else:
            print(f"  Status: WARNING - Translation pattern found but matching pattern differs")
            print(f"  Expected: {matching_pattern}")
            print(f"  Found: {found_pattern.get('matchingPattern')}")
            input("  Press Enter to acknowledge and continue...")
            return {}
    
    # Translation pattern does not exist, ask to create
    print(f"  Status: NOT FOUND - Translation pattern does not exist")
    create = input("  Create translation pattern? (Y/n): ").strip().lower()
    
    if create not in ['', 'y', 'yes']:
        print("\n  Translation pattern creation skipped.")
        print("  Please manually create the translation pattern in Control Hub.")
        input("  Press Enter to acknowledge and continue...")
        return {}
    
    # Clean up replacement pattern (remove dashes, spaces, keep only digits)
    replacement_clean = re.sub(r'[^0-9]', '', replacement_pattern)
    if len(replacement_clean) != 10:
        print(f"  Status: WARNING - Replacement pattern should be 10 digits, got {len(replacement_clean)} digits")
        print(f"  Original: {replacement_pattern}")
        print(f"  Cleaned: {replacement_clean}")
        proceed = input("  Proceed with cleaned value? (y/n): ").strip().lower()
        if proceed != 'y':
            print("  Translation pattern creation cancelled.")
            input("  Press Enter to continue...")
            return {}
    
    # Create translation pattern
    print(f"\n  Creating translation pattern...")
    create_data = {
        "name": translation_name,
        "matchingPattern": matching_pattern,
        "replacementPattern": replacement_clean
    }
    
    create_result = api.call("POST", "telephony/config/callRouting/translationPatterns",
                            data=create_data, params={"orgId": api.org_id})
    
    if "error" in create_result:
        print(f"  Status: FAILED - Error creating translation pattern: {create_result['error']}")
        print(f"  Please manually create the translation pattern in Control Hub.")
        input("  Press Enter to acknowledge and continue...")
        return {}
    
    pattern_id = create_result.get('id')
    print(f"  Status: SUCCESS - Translation pattern created")
    print(f"  Pattern ID: {pattern_id}")
    print(f"  Pattern Name: {translation_name}")
    
    input("\n  Press Enter to proceed to next step...")
    return {'id': pattern_id, 'name': translation_name}

def validate_call_park_extensions(api, location_data, filepath, read_excel_sheet, additional_tabs):
    """Validation 7: Validate call park extensions"""
    import re
    
    print(f"\nValidation 7: Validating call park extensions...")
    
    # Find location-specific tab
    location_name = location_data['name']
    location_tab = None
    
    for tab in additional_tabs:
        if tab.lower() == location_name.lower():
            location_tab = tab
            break
    
    if not location_tab:
        print(f"  Status: FAILED - Location tab '{location_name}' not found")
        input("  Press Enter to acknowledge and continue...")
        return {}
    
    # Read location tab
    location_data_sheet = read_excel_sheet(filepath, location_tab)
    if not location_data_sheet or len(location_data_sheet) < 46:
        print(f"  Status: FAILED - Could not read location tab or insufficient rows")
        input("  Press Enter to acknowledge and continue...")
        return {}
    
    # Extract call park data from B44, B45, C45 (column indices 1, 1, 2, rows 43, 44, 44)
    park_location = str(location_data_sheet[43][1]).strip() if len(location_data_sheet[43]) > 1 and location_data_sheet[43][1] else ""
    park_name_b45 = str(location_data_sheet[44][1]).strip() if len(location_data_sheet[44]) > 1 and location_data_sheet[44][1] else ""
    park_name_c45 = str(location_data_sheet[44][2]).strip() if len(location_data_sheet[44]) > 2 and location_data_sheet[44][2] else ""
    
    if not park_location or not park_name_b45 or not park_name_c45:
        print(f"  Status: WARNING - Call park data missing in location tab")
        print(f"    B44 (Location): {'[MISSING]' if not park_location else park_location}")
        print(f"    B45: {'[MISSING]' if not park_name_b45 else park_name_b45}")
        print(f"    C45: {'[MISSING]' if not park_name_c45 else park_name_c45}")
        input("  Press Enter to acknowledge and continue...")
        return {}
    
    # Verify location matches
    if park_location.lower() != location_name.lower():
        print(f"  Status: WARNING - Call park location '{park_location}' does not match inferred location '{location_name}'")
        input("  Press Enter to acknowledge and continue...")
        return {}
    
    # Combine B45 and C45 to get full range
    park_range = f"{park_name_b45} {park_name_c45}"
    print(f"  Call park range: {park_range}")
    
    # Parse range: "<Location> Park <ext#> thru <Location> Park <ext#>"
    # Extract start and end extension numbers
    match = re.search(r'Park (\d+) thru .* Park (\d+)', park_range)
    if not match:
        print(f"  Status: FAILED - Could not parse call park range format")
        print(f"  Expected format: '<Location> Park <ext#> thru <Location> Park <ext#>'")
        input("  Press Enter to acknowledge and continue...")
        return {}
    
    start_ext = int(match.group(1))
    end_ext = int(match.group(2))
    
    # Build list of required call park extensions
    required_parks = []
    for ext_num in range(start_ext, end_ext + 1):
        park_name = f"{location_name} Park {ext_num:02d}"
        park_ext = f"{ext_num:02d}"
        required_parks.append({
            'name': park_name,
            'extension': park_ext,
            'locationId': location_data['id']
        })
    
    print(f"  Required call park extensions: {len(required_parks)}")
    
    # Fetch existing call park extensions
    print(f"\n  Fetching existing call park extensions...")
    result = api.call("GET", "telephony/config/callParkExtensions",
                     params={"orgId": api.org_id, "locationId": location_data['id']})
    
    if "error" in result:
        print(f"  Status: FAILED - Error fetching call park extensions: {result['error']}")
        print(f"  Please manually check call park extensions in Control Hub.")
        input("  Press Enter to acknowledge and continue...")
        return {}
    
    existing_parks = result.get('callParkExtensions', [])
    print(f"  Found {len(existing_parks)} existing call park extensions")
    
    # Remove existing parks from required list
    to_create = []
    for park in required_parks:
        exists = False
        for existing in existing_parks:
            if existing.get('name') == park['name'] or existing.get('extension') == park['extension']:
                exists = True
                break
        if not exists:
            to_create.append(park)
    
    if not to_create:
        print(f"\n  Status: PASS - All required call park extensions already exist")
        input("  Press Enter to proceed to next step...")
        return {'created': 0}
    
    # Display table of call park extensions to create
    print(f"\n  Call park extensions to create:")
    print(f"  {'Name':<30} {'Extension':<10}")
    print(f"  {'-'*40}")
    for park in to_create:
        print(f"  {park['name']:<30} {park['extension']:<10}")
    
    create = input(f"\n  Create {len(to_create)} call park extension(s)? (Y/n): ").strip().lower()
    
    if create not in ['', 'y', 'yes']:
        print("\n  Call park extension creation skipped.")
        print("  Please manually create these call park extensions in Control Hub.")
        input("  Press Enter to acknowledge and continue...")
        return {'created': 0}
    
    # Create call park extensions
    print(f"\n  Creating call park extensions...")
    created_count = 0
    created_ids = []
    
    for park in to_create:
        print(f"\n  Creating '{park['name']}' (ext: {park['extension']})...")
        
        create_data = {
            "name": park['name'],
            "extension": park['extension']
        }
        
        create_result = api.call("POST", f"telephony/config/locations/{location_data['id']}/callParkExtensions",
                                data=create_data, params={"orgId": api.org_id})
        
        if "error" in create_result:
            print(f"    Status: FAILED - {create_result['error']}")
            print(f"    Please manually create this call park extension in Control Hub.")
        else:
            park_id = create_result.get('id')
            created_count += 1
            created_ids.append(park_id)
            print(f"    Status: SUCCESS - Call park extension created (ID: {park_id})")
    
    print(f"\n  Created {created_count} of {len(to_create)} call park extensions")
    
    if created_count < len(to_create):
        print(f"  Some call park extensions failed to create.")
        print(f"  Please manually check Control Hub.")
    
    proceed = input("\n  Press Enter to proceed to next step...")
    return {'created': created_count, 'ids': created_ids}
