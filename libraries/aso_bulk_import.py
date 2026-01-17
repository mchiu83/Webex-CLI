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

def validate_excel_file(filepath):
    """Validate Excel file structure and required tabs"""
    print(f"\nValidating Excel file: {filepath}")
    print(f"{'='*60}")
    
    required_tabs = ['Webex Users', 'Webex Side Cars', 'Webex Auto Attendant', 'Webex Hunt Groups']
    validation_results = []
    additional_tabs = []
    
    try:
        # Determine file type and load workbook
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
        
        # Validation 1: Check for required tabs
        print(f"\nValidation 1: Checking required tabs...")
        all_required_present = True
        
        for tab in required_tabs:
            if tab in sheet_names:
                print(f"  [{tab}] - Status: PASS")
                validation_results.append({'tab': tab, 'status': 'PASS'})
            else:
                print(f"  [{tab}] - Status: FAILED (Missing)")
                validation_results.append({'tab': tab, 'status': 'FAILED'})
                all_required_present = False
        
        if not all_required_present:
            print(f"\nOverall Status: FAILED - Missing required tabs")
            return False, None
        
        # Validation 2: Check for additional tabs
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

def aso_bulk_import_tool(api):
    """Main function for ASO Bulk Import Tool"""
    print("\n--- ASO Bulk Import Tool ---")
    
    # Check if bulk folder exists
    if not os.path.exists('bulk'):
        print("Status: FAILED - 'bulk' folder not found")
        print("Creating 'bulk' folder...")
        os.makedirs('bulk')
        print("Please place your 'aso_import' Excel file in the 'bulk' folder and try again.")
        return
    
    # Find aso_import file
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
    
    # Validate Excel file
    is_valid, additional_tabs = validate_excel_file(filepath)
    
    if not is_valid:
        print("\nValidation failed. Please fix the issues and try again.")
        return
    
    # Store additional tabs in memory
    if additional_tabs:
        print(f"\nAdditional tabs detected and cached:")
        for i, tab in enumerate(additional_tabs, 1):
            print(f"  {i}. {tab}")
    
    print("\nValidation complete. Ready for next steps.")
