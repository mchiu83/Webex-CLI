# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details


def view_license_usage(api):
    """
    Fetch and display license usage across all subscriptions
    for the organization using the Webex /licenses API.
    """
    print("\n" + "=" * 60)
    print("Organization License Usage")
    print("=" * 60)

    print("\nFetching licenses...")
    result = api.call("GET", "licenses", params={"orgId": api.org_id})

    if "error" in result:
        print(f"\nError fetching licenses: {result['error']}")
        return

    licenses = result.get("items", [])

    if not licenses:
        print("\nNo licenses found for this organization.")
        return

    # Group licenses by subscription
    subscriptions = {}
    for lic in licenses:
        sub_id = lic.get("subscriptionId", "Unknown")
        if sub_id not in subscriptions:
            subscriptions[sub_id] = []
        subscriptions[sub_id].append(lic)

    print(f"\nFound {len(licenses)} license(s) across {len(subscriptions)} subscription(s).\n")

    # Summary table header
    print(f"{'License Name':<45} {'Consumed':>10} {'Total':>10} {'Available':>10} {'Usage':>7}")
    print("-" * 85)

    total_consumed_all = 0
    total_units_all = 0

    for sub_id, sub_licenses in subscriptions.items():
        print(f"\n  Subscription: {sub_id}")
        print(f"  {'-' * 80}")

        for lic in sub_licenses:
            name = lic.get("name", "N/A")
            total_units = lic.get("totalUnits", 0)
            consumed = lic.get("consumedUnits", 0)
            available = total_units - consumed if total_units > 0 else 0
            usage_pct = (consumed / total_units * 100) if total_units > 0 else 0

            # Truncate long names
            display_name = name[:43] + ".." if len(name) > 45 else name

            usage_str = f"{usage_pct:.0f}%" if total_units > 0 else "N/A"
            total_str = str(total_units) if total_units > 0 else "Unlimited"
            avail_str = str(available) if total_units > 0 else "N/A"

            print(f"  {display_name:<45} {consumed:>10} {total_str:>10} {avail_str:>10} {usage_str:>7}")

            total_consumed_all += consumed
            if total_units > 0:
                total_units_all += total_units

    # Overall summary
    print(f"\n{'=' * 85}")
    overall_avail = total_units_all - total_consumed_all
    overall_pct = (total_consumed_all / total_units_all * 100) if total_units_all > 0 else 0
    print(f"  {'TOTAL':<45} {total_consumed_all:>10} {total_units_all:>10} {overall_avail:>10} {overall_pct:.0f}%")
    print(f"{'=' * 85}")
