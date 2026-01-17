def list_workspaces(api):
    print("\n--- List Workspaces ---")
    params = {"orgId": api.org_id}
    result = api.call("GET", "workspaces", params=params)
    
    if "error" in result:
        print(f"Error: {result['error']}")
        return None
    
    workspaces = result.get("items", [])
    if not workspaces:
        print("No workspaces found.")
        return None
    
    print(f"\nFound {len(workspaces)} workspace(s):")
    for i, ws in enumerate(workspaces, 1):
        print(f"{i}. {ws.get('displayName', 'N/A')} (ID: {ws.get('id', 'N/A')})")
    
    return workspaces
