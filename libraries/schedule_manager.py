# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

from datetime import datetime

VALID_SCHEDULES = ["24-7", "8-5NBD"]

def get_schedule_template(schedule_name):
    """Get schedule template based on name"""
    start_date = datetime.now().strftime("%Y-%m-%d")
    
    if schedule_name == "24-7":
        return {
            "type": "businessHours",
            "name": "24-7",
            "events": [
                {
                    "name": "Monday",
                    "startDate": start_date,
                    "endDate": start_date,
                    "allDayEnabled": True,
                    "recurrence": {"recurForEver": True, "recurWeekly": {"monday": True}}
                },
                {
                    "name": "Tuesday",
                    "startDate": start_date,
                    "endDate": start_date,
                    "allDayEnabled": True,
                    "recurrence": {"recurForEver": True, "recurWeekly": {"tuesday": True}}
                },
                {
                    "name": "Wednesday",
                    "startDate": start_date,
                    "endDate": start_date,
                    "allDayEnabled": True,
                    "recurrence": {"recurForEver": True, "recurWeekly": {"wednesday": True}}
                },
                {
                    "name": "Thursday",
                    "startDate": start_date,
                    "endDate": start_date,
                    "allDayEnabled": True,
                    "recurrence": {"recurForEver": True, "recurWeekly": {"thursday": True}}
                },
                {
                    "name": "Friday",
                    "startDate": start_date,
                    "endDate": start_date,
                    "allDayEnabled": True,
                    "recurrence": {"recurForEver": True, "recurWeekly": {"friday": True}}
                },
                {
                    "name": "Saturday",
                    "startDate": start_date,
                    "endDate": start_date,
                    "allDayEnabled": True,
                    "recurrence": {"recurForEver": True, "recurWeekly": {"saturday": True}}
                },
                {
                    "name": "Sunday",
                    "startDate": start_date,
                    "endDate": start_date,
                    "allDayEnabled": True,
                    "recurrence": {"recurForEver": True, "recurWeekly": {"sunday": True}}
                }
            ]
        }
    elif schedule_name == "8-5NBD":
        return {
            "type": "businessHours",
            "name": "8-5NBD",
            "events": [
                {
                    "name": "Monday",
                    "startDate": start_date,
                    "endDate": start_date,
                    "startTime": "08:00",
                    "endTime": "17:00",
                    "allDayEnabled": False,
                    "recurrence": {"recurForEver": True, "recurWeekly": {"monday": True}}
                },
                {
                    "name": "Tuesday",
                    "startDate": start_date,
                    "endDate": start_date,
                    "startTime": "08:00",
                    "endTime": "17:00",
                    "allDayEnabled": False,
                    "recurrence": {"recurForEver": True, "recurWeekly": {"tuesday": True}}
                },
                {
                    "name": "Wednesday",
                    "startDate": start_date,
                    "endDate": start_date,
                    "startTime": "08:00",
                    "endTime": "17:00",
                    "allDayEnabled": False,
                    "recurrence": {"recurForEver": True, "recurWeekly": {"wednesday": True}}
                },
                {
                    "name": "Thursday",
                    "startDate": start_date,
                    "endDate": start_date,
                    "startTime": "08:00",
                    "endTime": "17:00",
                    "allDayEnabled": False,
                    "recurrence": {"recurForEver": True, "recurWeekly": {"thursday": True}}
                },
                {
                    "name": "Friday",
                    "startDate": start_date,
                    "endDate": start_date,
                    "startTime": "08:00",
                    "endTime": "17:00",
                    "allDayEnabled": False,
                    "recurrence": {"recurForEver": True, "recurWeekly": {"friday": True}}
                }
            ]
        }
    return None

def validate_and_create_schedules(api, location_id, filepath):
    """Validate and create required schedules from Excel"""
    from libraries.aso_bulk_import import read_excel_sheet
    
    print(f"\n{'='*60}")
    print("Schedule Validation")
    print(f"{'='*60}")
    
    # Read Webex Auto Attendant sheet
    aa_data = read_excel_sheet(filepath, 'Webex Auto Attendant')
    if not aa_data or len(aa_data) < 30:
        print("  Error: Could not read Auto Attendant data")
        input("  Press Enter to continue...")
        return {}
    
    # Extract schedule names from J23-J29 (column index 9, rows 22-28)
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    schedule_map = {}
    errors = []
    
    for i, day in enumerate(days):
        row_idx = 22 + i
        if row_idx < len(aa_data) and len(aa_data[row_idx]) > 9:
            schedule_name = str(aa_data[row_idx][9]).strip() if aa_data[row_idx][9] else ""
            if not schedule_name:
                errors.append(f"{day}: Schedule name is missing")
            elif schedule_name not in VALID_SCHEDULES:
                errors.append(f"{day}: Invalid schedule '{schedule_name}' (must be '24-7' or '8-5NBD')")
            else:
                schedule_map[day] = schedule_name
        else:
            errors.append(f"{day}: Row not found in Excel")
    
    if errors:
        print("\n  Schedule Validation Errors:")
        for error in errors:
            print(f"    - {error}")
        print("\n  Please fix schedule names in Excel (cells J23-J29)")
        input("  Press Enter to acknowledge and continue...")
        return {}
    
    # Display schedule table
    print("\n  Required Schedules:")
    print(f"  {'Day':<12} {'Schedule':<10}")
    print(f"  {'-'*25}")
    for day, schedule in schedule_map.items():
        print(f"  {day:<12} {schedule:<10}")
    
    # Get unique schedules
    unique_schedules = set(schedule_map.values())
    print(f"\n  Unique schedules needed: {', '.join(unique_schedules)}")
    
    # Fetch existing schedules
    print(f"\n  Fetching existing schedules from location...")
    result = api.call("GET", f"telephony/config/locations/{location_id}/schedules", 
                     params={"orgId": api.org_id})
    
    if "error" in result:
        print(f"  Error fetching schedules: {result['error']}")
        input("  Press Enter to continue...")
        return {}
    
    existing_schedules = {s['name']: s['id'] for s in result.get('schedules', [])}
    schedule_ids = {}
    
    # Check which schedules exist
    missing_schedules = []
    for schedule_name in unique_schedules:
        if schedule_name in existing_schedules:
            schedule_ids[schedule_name] = existing_schedules[schedule_name]
            print(f"  [OK] Schedule '{schedule_name}' exists (ID: {existing_schedules[schedule_name]})")
        else:
            missing_schedules.append(schedule_name)
            print(f"  [MISSING] Schedule '{schedule_name}' does not exist")
    
    if not missing_schedules:
        print(f"\n  All required schedules exist!")
        return schedule_ids
    
    # Ask permission to create missing schedules
    print(f"\n  Missing schedules: {', '.join(missing_schedules)}")
    create = input("  Create missing schedules? (Y/n): ").strip().lower()
    
    if create not in ['', 'y', 'yes']:
        print("\n  Schedule creation skipped.")
        print("  Please manually create these schedules in Control Hub before Auto Attendant creation.")
        input("  Press Enter to acknowledge and continue...")
        return schedule_ids
    
    # Create missing schedules
    for schedule_name in missing_schedules:
        print(f"\n  Creating schedule '{schedule_name}'...")
        
        schedule_data = get_schedule_template(schedule_name)
        if not schedule_data:
            print(f"    Error: No template for '{schedule_name}'")
            continue
        
        result = api.call("POST", f"telephony/config/locations/{location_id}/schedules",
                         data=schedule_data, params={"orgId": api.org_id})
        
        if "error" in result:
            print(f"    Error: {result['error']}")
            print(f"    Please manually create '{schedule_name}' in Control Hub")
        else:
            schedule_id = result.get('id')
            schedule_ids[schedule_name] = schedule_id
            print(f"    Success: Schedule created (ID: {schedule_id})")
    
    if len(schedule_ids) < len(unique_schedules):
        print("\n  Some schedules failed to create.")
        print("  Please manually create missing schedules in Control Hub.")
    
    input("\n  Press Enter to continue...")
    return schedule_ids
