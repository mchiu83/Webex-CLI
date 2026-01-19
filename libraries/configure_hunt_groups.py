# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

def configure_hunt_groups(api, location_data, workspace_map, data_rows, filepath):
    """Configure hunt groups from Webex Hunt Groups sheet"""
    from libraries.aso_bulk_import import read_excel_sheet
    
    print(f"\n{'='*60}")
    proceed = input("\nProceed with hunt group configuration? (Y/n): ").strip().lower()
    if proceed not in ['', 'y', 'yes']:
        print("\nHunt group configuration skipped.")
        return
    
    # Read Webex Hunt Groups sheet
    huntgroup_data = read_excel_sheet(filepath, 'Webex Hunt Groups')
    if not huntgroup_data or len(huntgroup_data) < 7:
        print("  Skipped: No hunt group data found")
        return
    
    # Get location timezone
    location_result = api.call("GET", f"locations/{location_data['id']}", params={"orgId": api.org_id})
    if "error" in location_result:
        print(f"  Error fetching location details: {location_result['error']}")
        return
    
    timezone = location_result.get('timeZone', 'America/Chicago')
    
    # Build extension to workspace ID map
    ext_to_workspace = {}
    for row_idx, workspace_id in workspace_map.items():
        row = data_rows[row_idx - 2]
        extension = str(row[4]).strip() if len(row) > 4 else ""
        if extension:
            ext_to_workspace[extension] = workspace_id
    
    # Parse hunt groups (rows 4-6, 7-9, etc.)
    hunt_groups = []
    row_idx = 3  # Start at row 4 (0-indexed)
    
    while row_idx < len(huntgroup_data):
        # Check if we have a hunt group (name in column A)
        if len(huntgroup_data[row_idx]) > 0 and huntgroup_data[row_idx][0]:
            hg_name = str(huntgroup_data[row_idx][0]).strip()
            hg_phone_raw = str(huntgroup_data[row_idx][1]).strip() if len(huntgroup_data[row_idx]) > 1 and huntgroup_data[row_idx][1] else ""
            hg_phone = None if not hg_phone_raw or hg_phone_raw.upper() == 'N/A' else hg_phone_raw
            hg_ext = str(huntgroup_data[row_idx][2]).strip() if len(huntgroup_data[row_idx]) > 2 else ""
            
            # Collect agent extensions from 3 rows
            agent_extensions = []
            for i in range(3):
                if row_idx + i < len(huntgroup_data) and len(huntgroup_data[row_idx + i]) > 3:
                    ext = str(huntgroup_data[row_idx + i][3]).strip() if huntgroup_data[row_idx + i][3] else None
                    if ext:
                        agent_extensions.append(ext)
            
            # Get policy and other settings from first row
            policy = str(huntgroup_data[row_idx][5]).strip().upper() if len(huntgroup_data[row_idx]) > 5 and huntgroup_data[row_idx][5] else "REGULAR"
            next_agent_rings = int(float(huntgroup_data[row_idx][6])) if len(huntgroup_data[row_idx]) > 6 and huntgroup_data[row_idx][6] else 3
            
            hunt_groups.append({
                'name': hg_name,
                'phoneNumber': hg_phone,
                'extension': hg_ext,
                'agent_extensions': agent_extensions,
                'policy': policy,
                'nextAgentRings': next_agent_rings
            })
        
        row_idx += 3  # Move to next hunt group
    
    if not hunt_groups:
        print("  Skipped: No hunt groups found")
        return
    
    # Process each hunt group
    for hg in hunt_groups:
        print(f"\n{'='*60}")
        print(f"Hunt Group: {hg['name']}")
        print(f"{'='*60}")
        
        # Map agent extensions to workspace IDs
        agent_ids = []
        for ext in hg['agent_extensions']:
            if ext in ext_to_workspace:
                agent_ids.append({"id": ext_to_workspace[ext]})
            else:
                print(f"  Warning: Extension {ext} not found in created workspaces")
        
        if not agent_ids:
            print(f"  Skipped: No valid agents found")
            continue
        
        # Remove trailing number from name for customName
        custom_name = hg['name']
        if custom_name and custom_name[-1].isdigit():
            # Remove trailing digits and spaces
            custom_name = custom_name.rstrip('0123456789').strip()
        
        # Build hunt group data
        hg_data = {
            "name": hg['name'],
            "extension": int(hg['extension']),
            "timeZone": timezone,
            "callPolicies": {
                "policy": hg['policy'],
                "waitingEnabled": False,
                "noAnswer": {
                    "nextAgentEnabled": True,
                    "nextAgentRings": hg['nextAgentRings']
                }
            },
            "agents": agent_ids,
            "enabled": True,
            "huntGroupCallerIdForOutgoingCallsEnabled": True,
            "directLineCallerIdName": {
                "selection": "CUSTOM_NAME",
                "customName": custom_name
            },
            "dialByName": custom_name
        }
        
        if hg['phoneNumber']:
            hg_data["phoneNumber"] = int(hg['phoneNumber'])
        
        # Display configuration table
        print(f"\nConfiguration:")
        print(f"  Name: {hg_data['name']}")
        print(f"  Extension: {hg_data['extension']}")
        if hg['phoneNumber']:
            print(f"  Phone Number: {hg['phoneNumber']}")
        print(f"  Time Zone: {timezone}")
        print(f"  Policy: {hg_data['callPolicies']['policy']}")
        print(f"  Next Agent Rings: {hg_data['callPolicies']['noAnswer']['nextAgentRings']}")
        print(f"  Agents: {len(agent_ids)}")
        print(f"  Custom Name: {custom_name}")
        
        # Confirm
        confirm = input("\nProceed with this hunt group? (Y/n): ").strip().lower()
        if confirm not in ['', 'y', 'yes']:
            print("  Skipped")
            continue
        
        # Allow modifications
        modify = input("Modify any attributes? (y/N): ").strip().lower()
        if modify == 'y':
            hg_data['name'] = input(f"  Name [{hg_data['name']}]: ").strip() or hg_data['name']
            ext_input = input(f"  Extension [{hg_data['extension']}]: ").strip()
            if ext_input:
                hg_data['extension'] = int(ext_input)
            if 'phoneNumber' in hg_data:
                new_phone = input(f"  Phone Number [{hg_data['phoneNumber']}]: ").strip()
                if new_phone:
                    hg_data['phoneNumber'] = int(new_phone)
            policy_input = input(f"  Policy [{hg_data['callPolicies']['policy']}]: ").strip().upper()
            if policy_input:
                hg_data['callPolicies']['policy'] = policy_input
            rings_input = input(f"  Next Agent Rings [{hg_data['callPolicies']['noAnswer']['nextAgentRings']}]: ").strip()
            if rings_input:
                hg_data['callPolicies']['noAnswer']['nextAgentRings'] = int(rings_input)
        
        # Create hunt group
        print(f"\n  Creating hunt group...")
        result = api.call("POST", f"telephony/config/locations/{location_data['id']}/huntGroups",
                         data=hg_data, params={"orgId": api.org_id})
        
        if "error" in result:
            print(f"  Error: {result['error']}")
            print(f"  Please check manually in Control Hub")
        else:
            print(f"  Success: Hunt group created successfully!")
