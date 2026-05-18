# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

import os
import re
import io
import struct
import wave
import audioop
import requests

# ---------------------------------------------------------------------------
# Timezone display name -> IANA mapping
# ---------------------------------------------------------------------------
TIMEZONE_MAP = {
    'eastern':    'America/New_York',
    'central':    'America/Chicago',
    'mountain':   'America/Denver',
    'pacific':    'America/Los_Angeles',
    'alaska':     'America/Anchorage',
    'hawaii':     'Pacific/Honolulu',
    'arizona':    'America/Phoenix',
    'atlantic':   'America/Halifax',
}

# ---------------------------------------------------------------------------
# Sheet layout constants (0-indexed columns)
# ---------------------------------------------------------------------------
COL_KEY        = 1   # B  - key digit (0-9, *, #)
COL_ACTION     = 2   # C  - action string
COL_DEST       = 3   # D  - destination number / extension
COL_LABEL      = 4   # E  - label / display name
COL_SCHED_DAY  = 8   # I  - day name (business hours section)
COL_SCHED_NAME = 9   # J  - schedule name

# Row offsets within each 23-row AA block (0-indexed relative to block start)
OFFSET_HEADER      = 0   # AA name, DID
OFFSET_KEYS_START  = 1   # key 0
OFFSET_KEYS_END    = 11  # key # (rows +1 to +11)
OFFSET_SCHED_START = 12  # Monday  (rows +12 to +18)
OFFSET_SCHED_END   = 18  # Sunday
OFFSET_TIMEZONE    = 20  # TimeZone row

# The four AA blocks start at these 0-indexed row positions
AA_BLOCK_STARTS = [10, 33, 58, 83]   # rows 11, 34, 59, 84 in 1-indexed

# Greeting audio file names are in column M (index 12), rows 4-8 (0-indexed 3-7)
GREETING_ROWS = [4, 5, 6, 7]   # 0-indexed rows that contain menu -> wav mappings
COL_MENU_NAME = 11  # L - menu name
COL_WAV_NAME  = 12  # M - wav file name


def _resolve_timezone(raw: str) -> str:
    """Convert display timezone name to IANA format."""
    key = raw.strip().lower()
    return TIMEZONE_MAP.get(key, raw.strip())


def _clean_ext(value: str) -> str:
    """Strip non-digit characters and return as string."""
    return re.sub(r'[^0-9]', '', str(value).strip())


def _parse_aa_blocks(sheet_data: list, fallback_timezone: str = "America/Chicago") -> list:
    """
    Parse the four fixed AA blocks from the sheet.
    Returns a list of dicts, one per AA.
    """
    aas = []

    for block_start in AA_BLOCK_STARTS:
        if block_start >= len(sheet_data):
            break

        header_row = sheet_data[block_start]
        aa_name = str(header_row[COL_KEY]).strip() if len(header_row) > COL_KEY and header_row[COL_KEY] else ""
        if not aa_name:
            continue

        # DID / extension from col D of header row
        # Values prefixed with "DID/Ext:" are AA extensions.
        # Raw 10-digit numbers are store DIDs to assign as phoneNumber.
        did_raw = str(header_row[COL_DEST]).strip() if len(header_row) > COL_DEST and header_row[COL_DEST] else ""
        if re.match(r'(?i)did/ext:', did_raw):
            did_raw   = re.sub(r'(?i)did/ext:\s*', '', did_raw).strip()
            did_clean = _clean_ext(did_raw)
        else:
            did_clean = _clean_ext(did_raw)   # raw DID — digits only

        # Key configurations
        key_configs = []
        for offset in range(OFFSET_KEYS_START, OFFSET_KEYS_END + 1):
            row_idx = block_start + offset
            if row_idx >= len(sheet_data):
                break
            row = sheet_data[row_idx]
            key_val  = str(row[COL_KEY]).strip()    if len(row) > COL_KEY    and row[COL_KEY]    else ""
            action   = str(row[COL_ACTION]).strip() if len(row) > COL_ACTION and row[COL_ACTION] else ""
            dest     = str(row[COL_DEST]).strip()   if len(row) > COL_DEST   and row[COL_DEST]   else ""
            label    = str(row[COL_LABEL]).strip()  if len(row) > COL_LABEL  and row[COL_LABEL]  else ""

            if not key_val or not action:
                continue

            key_configs.append({
                'key':    key_val,
                'action': action,
                'dest':   dest,
                'label':  label,
            })

        # Business hours schedule (day -> schedule name)
        schedule_map = {}
        for offset in range(OFFSET_SCHED_START, OFFSET_SCHED_END + 1):
            row_idx = block_start + offset
            if row_idx >= len(sheet_data):
                break
            row = sheet_data[row_idx]
            day   = str(row[COL_SCHED_DAY]).strip()  if len(row) > COL_SCHED_DAY  and row[COL_SCHED_DAY]  else ""
            sched = str(row[COL_SCHED_NAME]).strip() if len(row) > COL_SCHED_NAME and row[COL_SCHED_NAME] else ""
            if day and sched:
                schedule_map[day] = sched

        # Timezone
        tz_raw = ""
        tz_row_idx = block_start + OFFSET_TIMEZONE
        if tz_row_idx < len(sheet_data):
            tz_row = sheet_data[tz_row_idx]
            if len(tz_row) > COL_SCHED_NAME and tz_row[COL_SCHED_NAME]:
                tz_raw = str(tz_row[COL_SCHED_NAME]).strip()

        aas.append({
            'name':         aa_name,
            'did':          did_clean,
            'key_configs':  key_configs,
            'schedule_map': schedule_map,
            'timezone_raw': tz_raw,
            'timezone':     _resolve_timezone(tz_raw) if tz_raw else fallback_timezone,
        })

    return aas


def _parse_greeting_map(sheet_data: list) -> dict:
    """
    Read the menu -> wav filename mapping from the top of the sheet (rows 4-8, 0-indexed 3-7).
    Returns dict: { menu_name_lower: wav_filename }
    """
    greeting_map = {}
    for row_idx in GREETING_ROWS:
        if row_idx >= len(sheet_data):
            break
        row = sheet_data[row_idx]
        menu_name = str(row[COL_MENU_NAME]).strip() if len(row) > COL_MENU_NAME and row[COL_MENU_NAME] else ""
        wav_name  = str(row[COL_WAV_NAME]).strip()  if len(row) > COL_WAV_NAME  and row[COL_WAV_NAME]  else ""
        if menu_name and wav_name:
            greeting_map[menu_name.lower()] = wav_name
    return greeting_map


def _get_location_announcements(api, location_id: str) -> dict:
    """
    Fetch existing announcement files at the location level.
    Returns dict: { filename_lower: announcement_id }
    """
    result = api.call(
        "GET",
        "telephony/config/announcements",
        params={"orgId": api.org_id, "locationId": location_id}
    )
    if "error" in result:
        return {}
    out = {}
    for a in result.get("announcements", []):
        ann_id = a.get("id")
        if not ann_id or str(a.get("level", "")).upper() != "LOCATION":
            continue
        if a.get("fileName"):
            out[a["fileName"].lower()] = ann_id
    return out


def _get_org_announcements(api) -> dict:
    """
    Fetch existing announcement files at the organization (global) level.
    Returns dict: { filename_lower: announcement_id }
    """
    result = api.call(
        "GET",
        "telephony/config/announcements",
        params={"orgId": api.org_id}
    )
    if "error" in result:
        return {}
    out = {}
    for a in result.get("announcements", []):
        ann_id = a.get("id")
        if not ann_id:
            continue
        if a.get("fileName"):
            out[a["fileName"].lower()] = ann_id
    return out


def _convert_to_ulaw_8k_mono(input_path: str) -> io.BytesIO:
    """
    Convert any PCM wav file to mono 8kHz 8-bit u-law in memory.
    Webex requires: mono, 8000 Hz, u-law (CCITT G.711), WAV format.
    Returns a BytesIO buffer ready for upload.
    """
    with wave.open(input_path, 'rb') as src:
        n_channels = src.getnchannels()
        samp_width = src.getsampwidth()
        frame_rate = src.getframerate()
        pcm_data   = src.readframes(src.getnframes())

    # Normalise to 16-bit signed PCM
    if samp_width == 1:
        pcm_data   = audioop.bias(pcm_data, 1, -128)
        pcm_data   = audioop.lin2lin(pcm_data, 1, 2)
        samp_width = 2
    elif samp_width != 2:
        pcm_data   = audioop.lin2lin(pcm_data, samp_width, 2)
        samp_width = 2

    # Stereo -> mono
    if n_channels == 2:
        pcm_data   = audioop.tomono(pcm_data, 2, 0.5, 0.5)
        n_channels = 1

    # Resample to 8000 Hz
    if frame_rate != 8000:
        pcm_data, _ = audioop.ratecv(pcm_data, 2, 1, frame_rate, 8000, None)

    # PCM16 -> u-law
    ulaw_data = audioop.lin2ulaw(pcm_data, 2)

    # Build WAV container with MULAW (format 7) header manually —
    # Python's wave module only writes PCM.
    n_frames        = len(ulaw_data)
    byte_rate       = 8000
    block_align     = 1
    bits_per_sample = 8
    chunk_size      = 36 + 2 + n_frames   # 36 standard + 2 extra bytes in fmt + data

    buf = io.BytesIO()
    buf.write(struct.pack('<4sI4s', b'RIFF', chunk_size, b'WAVE'))
    buf.write(struct.pack('<4sI',   b'fmt ', 18))
    buf.write(struct.pack('<HHIIHH',
        7, 1, 8000, byte_rate, block_align, bits_per_sample))
    buf.write(struct.pack('<H', 0))          # cbSize (extra param bytes = 0)
    buf.write(struct.pack('<4sI', b'data', n_frames))
    buf.write(ulaw_data)
    buf.seek(0)
    return buf


def _upload_announcement(api, location_id: str, wav_path: str, wav_filename: str,
                          scope: str = "location") -> str | None:
    """
    Convert and upload a wav file to the announcement repository.
    scope: "location" uploads to the specific location; "global" uploads org-wide.
    Automatically converts to mono 8kHz u-law as required by Webex.
    Returns the new announcement ID, or None on failure.
    """
    if scope == "global":
        url = f"{api.base_url}/telephony/config/announcements"
    else:
        url = f"{api.base_url}/telephony/config/locations/{location_id}/announcements"

    headers = {"Authorization": f"Bearer {api.token}"}
    params  = {"orgId": api.org_id}
    name    = wav_filename.rsplit('.', 1)[0]

    try:
        print(f"    Converting to mono 8kHz u-law...")
        converted = _convert_to_ulaw_8k_mono(wav_path)
        print(f"    Uploading ({len(converted.getvalue())} bytes)...")

        files = {"file": (wav_filename, converted, "audio/wav")}
        data  = {"name": name}
        response = requests.post(url, headers=headers, params=params, files=files, data=data)
        api.api_logger.info(f"Upload announcement ({scope}): POST {url} -> {response.status_code}")
        api.api_logger.info(f"Response: {response.text}")

        if response.status_code in [200, 201]:
            return response.json().get("id")
        else:
            api.api_logger.error(f"Upload failed: {response.status_code} - {response.text}")
            print(f"    API error {response.status_code}: {response.text[:200]}")
            return None
    except Exception as e:
        api.api_logger.error(f"Upload exception: {e}")
        print(f"    Exception during upload: {e}")
        return None


def _resolve_announcements(api, location_id: str, greeting_map: dict) -> dict | None:
    """
    For each wav file referenced in greeting_map, ensure it exists in Control Hub.
    Checks location-level first, then org-level (global).
    If not found in either, checks /bulk folder and offers to upload with a
    choice of Location or Global scope.
    Returns dict: { wav_filename_lower: announcement_id } or None if a required file is missing.
    """
    location_existing = _get_location_announcements(api, location_id)
    org_existing      = _get_org_announcements(api)
    resolved = {}   # wav_lower -> announcement_id
    levels   = {}   # wav_lower -> "LOCATION" | "ORGANIZATION"
    errors   = []

    for menu_name, wav_filename in greeting_map.items():
        wav_lower = wav_filename.lower()

        # Check location-level first
        if wav_lower in location_existing:
            print(f"  [{wav_filename}] - Found in Control Hub at location level (ID: {location_existing[wav_lower]})")
            resolved[wav_lower] = location_existing[wav_lower]
            levels[wav_lower]   = "LOCATION"
            continue

        # Check org/global level
        if wav_lower in org_existing:
            print(f"  [{wav_filename}] - Found in Control Hub at global level (ID: {org_existing[wav_lower]})")
            resolved[wav_lower] = org_existing[wav_lower]
            levels[wav_lower]   = "ORGANIZATION"
            continue

        # Not in Control Hub at either level — check /bulk folder
        bulk_path = os.path.join("bulk", wav_filename)
        if os.path.exists(bulk_path):
            print(f"  [{wav_filename}] - Not in Control Hub, found in /bulk folder")
            upload = input(f"    Upload '{wav_filename}' to Control Hub? (Y/n): ").strip().lower()
            if upload not in ['', 'y', 'yes']:
                print(f"    Skipped upload.")
                errors.append(wav_filename)
                continue

            # Ask scope
            print(f"    Upload scope:")
            print(f"      L - Location only (this location)")
            print(f"      G - Global (available to all locations)")
            scope_choice = input(f"    Choice (L/G): ").strip().upper()
            scope = "global" if scope_choice == "G" else "location"
            scope_label = "globally" if scope == "global" else f"to location"

            print(f"    Uploading {scope_label}...")
            ann_id = _upload_announcement(api, location_id, bulk_path, wav_filename, scope)
            if ann_id:
                print(f"    Uploaded successfully (ID: {ann_id})")
                resolved[wav_lower] = ann_id
                levels[wav_lower]   = "ORGANIZATION" if scope == "global" else "LOCATION"
            elif ann_id is None:
                # Upload may have failed because file already exists under a different label.
                # Re-fetch both repositories and try to match by fileName.
                print(f"    Checking if file already exists under a different label...")
                refreshed_loc = _get_location_announcements(api, location_id)
                refreshed_org = _get_org_announcements(api)
                if wav_lower in refreshed_loc:
                    print(f"    Found at location level after re-check (ID: {refreshed_loc[wav_lower]})")
                    resolved[wav_lower] = refreshed_loc[wav_lower]
                    levels[wav_lower]   = "LOCATION"
                elif wav_lower in refreshed_org:
                    print(f"    Found at global level after re-check (ID: {refreshed_org[wav_lower]})")
                    resolved[wav_lower] = refreshed_org[wav_lower]
                    levels[wav_lower]   = "ORGANIZATION"
                else:
                    print(f"    File not found after re-check. Please delete the duplicate in Control Hub and retry.")
                    errors.append(wav_filename)
        else:
            print(f"  [{wav_filename}] - NOT found in Control Hub or /bulk folder")
            errors.append(wav_filename)

    if errors:
        print(f"\n  The following required audio files are missing:")
        for f in errors:
            print(f"    - {f}")
        print(f"  Please add the missing files to the /bulk folder or upload them to Control Hub manually.")
        return None

    return resolved, levels


def _build_menu_greeting(wav_filename: str, announcement_id: str, level: str = "ORGANIZATION") -> dict:
    """Build the greeting block for a menu using a custom audio file."""
    return {
        "greeting": "CUSTOM",
        "audioAnnouncementFile": {
            "id":            announcement_id,
            "fileName":      wav_filename,
            "mediaFileType": "WAV",
            "level":         level
        }
    }


def _map_action(action_str: str, dest: str, label: str, aa_ext_map: dict) -> dict | None:
    """
    Convert a spreadsheet action string into a Webex API keyConfiguration dict.
    Returns None for unsupported / empty actions.
    """
    action_lower = action_str.strip().lower()
    dest_clean   = _clean_ext(dest)

    if action_lower == "repeat menu":
        return {"action": "REPEAT_MENU"}

    if action_lower in ("transfer to number", "transfer to sub menu"):
        if not dest_clean:
            return None
        return {
            "action": "TRANSFER_WITHOUT_PROMPT",
            "value":  dest_clean,
        }

    # "Transfer to Recording" and anything else — skip
    return None


def _validate_schedule_exists(api, location_id: str, schedule_name: str) -> bool:
    """Confirm a schedule with the given name exists at the location."""
    result = api.call(
        "GET",
        f"telephony/config/locations/{location_id}/schedules",
        params={"orgId": api.org_id}
    )
    if "error" in result:
        return False
    return any(
        s.get("name", "").strip().lower() == schedule_name.strip().lower()
        for s in result.get("schedules", [])
    )


def _build_aa_payload(aa: dict, greeting_block: dict, schedule_name: str, location_id: str) -> dict:
    """
    Build the full POST body for creating an Auto Attendant.
    After-hours menu mirrors business hours menu (same keys, same greeting).
    businessSchedule is a plain schedule name string as required by the API.
    firstName/lastName are derived from the AA name (required fields).
    """
    key_configurations = []
    for kc in aa['key_configs']:
        mapped = _map_action(kc['action'], kc['dest'], kc['label'], {})
        if mapped:
            mapped['key'] = kc['key']
            key_configurations.append(mapped)

    menu = {
        **greeting_block,
        "extensionEnabled": False,
        "keyConfigurations": key_configurations,
    }

    # Derive firstName / lastName from the AA name (required by API)
    name_parts = aa['name'].split('-', 2)
    first_name = name_parts[0].strip() if len(name_parts) > 0 else aa['name']
    last_name  = name_parts[-1].strip() if len(name_parts) > 1 else "AA"

    payload = {
        "name":             aa['name'],
        "firstName":        first_name,
        "lastName":         last_name,
        "timeZone":         aa['timezone'],
        "businessHoursMenu": menu,
        "afterHoursMenu":    menu,
        "extensionDialing":  "GROUP",
        "nameDialing":       "GROUP",
        "languageCode":      "en_us",
    }

    if aa['did']:
        if len(aa['did']) >= 10:
            payload["phoneNumber"] = f"+1{aa['did']}" if len(aa['did']) == 10 else aa['did']
        else:
            payload["extension"] = aa['did']

    if schedule_name:
        payload["businessSchedule"] = schedule_name

    return payload


def _check_existing_aa(api, location_id: str, aa_name: str) -> str | None:
    """Return the ID of an existing AA with the same name, or None."""
    result = api.call(
        "GET",
        "telephony/config/autoAttendants",
        params={"orgId": api.org_id, "locationId": location_id, "name": aa_name}
    )
    if "error" in result:
        return None
    for aa in result.get("autoAttendants", []):
        if aa.get("name", "").strip().lower() == aa_name.strip().lower():
            return aa.get("id")
    return None


def _get_aa_details(api, location_id: str, aa_id: str) -> dict | None:
    """Fetch full details for an existing AA."""
    result = api.call(
        "GET",
        f"telephony/config/locations/{location_id}/autoAttendants/{aa_id}",
        params={"orgId": api.org_id}
    )
    return None if "error" in result else result


def _compare_key_configs(desired: list, actual_menu: dict) -> list:
    """
    Compare desired key configs (from sheet) against the actual menu from Control Hub.
    Returns a list of diff strings describing mismatches.
    """
    diffs = []
    actual_keys = {
        kc.get("key"): kc
        for kc in actual_menu.get("keyConfigurations", [])
    }

    for kc in desired:
        action_lower = kc['action'].strip().lower()
        dest_clean   = _clean_ext(kc['dest'])
        key          = kc['key']

        # Determine what we expect in the API
        if action_lower == "repeat menu":
            expected_action = "REPEAT_MENU"
            expected_dest   = None
            expected_label  = None
        elif action_lower in ("transfer to number", "transfer to sub menu"):
            if not dest_clean:
                continue
            expected_action = "TRANSFER_WITHOUT_PROMPT"
            expected_dest   = dest_clean
            expected_label  = kc['label'] or dest_clean
        else:
            continue  # unsupported action, skip

        actual_kc = actual_keys.get(key)
        if not actual_kc:
            diffs.append(f"  Key {key}: MISSING in Control Hub (expected {expected_action}"
                         + (f" -> {expected_dest}" if expected_dest else "") + ")")
            continue

        actual_action = actual_kc.get("action", "")
        actual_dest   = _clean_ext(str(actual_kc.get("destination", "")))
        actual_label  = actual_kc.get("name", "")

        if actual_action != expected_action:
            diffs.append(f"  Key {key}: action  expected='{expected_action}'  actual='{actual_action}'")
        if expected_dest and actual_dest != expected_dest:
            diffs.append(f"  Key {key}: dest    expected='{expected_dest}'  actual='{actual_dest}'")
        if expected_label and actual_label != expected_label:
            diffs.append(f"  Key {key}: label   expected='{expected_label}'  actual='{actual_label}'")

    # Check for extra keys in Control Hub not in the sheet
    desired_keys = {kc['key'] for kc in desired if kc.get('action', '').strip().lower() not in ('', 'transfer to recording')}
    for key in actual_keys:
        if key not in desired_keys:
            diffs.append(f"  Key {key}: present in Control Hub but not in worksheet")

    return diffs


def _validate_and_update_existing_aa(
    api, location_id: str, aa_id: str, aa: dict,
    greeting_block: dict, schedule_name: str
) -> bool:
    """
    Fetch the existing AA, compare all settings against the worksheet definition,
    print a diff, and offer to apply corrections.
    Returns True if the AA is ready to use (either matched or updated), False on error.
    """
    print(f"  Fetching existing AA details...")
    details = _get_aa_details(api, location_id, aa_id)
    if not details:
        print(f"  Warning: Could not fetch details for existing AA. Treating as unverified.")
        return True

    diffs = []

    # --- Extension / phone number ---
    actual_ext   = str(details.get("extension", "")).strip()
    actual_phone = _clean_ext(str(details.get("phoneNumber", "")).strip())
    if aa['did']:
        if len(aa['did']) >= 10:
            expected_phone = f"+1{aa['did']}" if len(aa['did']) == 10 else aa['did']
            if actual_phone != _clean_ext(expected_phone):
                diffs.append(f"  Phone number: expected='{expected_phone}'  actual='{details.get('phoneNumber', '')}'")
        else:
            if actual_ext != aa['did']:
                diffs.append(f"  Extension:    expected='{aa['did']}'  actual='{actual_ext}'")

    # --- Timezone ---
    actual_tz = details.get("timeZone", "")
    if actual_tz != aa['timezone']:
        diffs.append(f"  Timezone:     expected='{aa['timezone']}'  actual='{actual_tz}'")

    # --- Business hours schedule ---
    actual_sched_name = details.get("businessSchedule", "")
    if schedule_name and aa['schedule_map']:
        expected_sched_name = next(iter(aa['schedule_map'].values()))
        if actual_sched_name.strip().lower() != expected_sched_name.strip().lower():
            diffs.append(f"  Schedule:     expected='{expected_sched_name}'  actual='{actual_sched_name}'")

    # --- Business hours menu greeting ---
    bh_menu = details.get("businessHoursMenu", {})
    actual_greeting = bh_menu.get("greeting", "")
    expected_greeting = greeting_block.get("greeting", "DEFAULT")
    if actual_greeting != expected_greeting:
        diffs.append(f"  Greeting:     expected='{expected_greeting}'  actual='{actual_greeting}'")
    elif expected_greeting == "CUSTOM":
        actual_wav = bh_menu.get("audioAnnouncementFile", {}).get("fileName", "")
        expected_wav = greeting_block.get("audioAnnouncementFile", {}).get("fileName", "")
        if actual_wav.lower() != expected_wav.lower():
            diffs.append(f"  Audio file:   expected='{expected_wav}'  actual='{actual_wav}'")

    # --- Key configurations ---
    key_diffs = _compare_key_configs(aa['key_configs'], bh_menu)
    diffs.extend(key_diffs)

    # --- Report ---
    if not diffs:
        print(f"  Status: MATCH — all settings are correct, no changes needed.")
        return True

    print(f"\n  Status: MISMATCH — {len(diffs)} difference(s) found:")
    for d in diffs:
        print(d)

    apply = input(f"\n  Apply corrections to '{aa['name']}'? (Y/n): ").strip().lower()
    if apply not in ['', 'y', 'yes']:
        print(f"  Skipped: '{aa['name']}' left unchanged.")
        return True

    # Build corrected key configurations
    key_configurations = []
    for kc in aa['key_configs']:
        action_lower = kc['action'].strip().lower()
        dest_clean   = _clean_ext(kc['dest'])
        if action_lower == "repeat menu":
            key_configurations.append({"key": kc['key'], "action": "REPEAT_MENU"})
        elif action_lower in ("transfer to number", "transfer to sub menu") and dest_clean:
            key_configurations.append({
                "key":    kc['key'],
                "action": "TRANSFER_WITHOUT_PROMPT",
                "value":  dest_clean,
            })

    menu = {
        **greeting_block,
        "extensionEnabled":  False,
        "keyConfigurations": key_configurations,
    }

    update_payload = {
        "name":              aa['name'],
        "timeZone":          aa['timezone'],
        "businessHoursMenu": menu,
        "afterHoursMenu":    menu,
    }

    if aa['did']:
        if len(aa['did']) >= 10:
            update_payload["phoneNumber"] = f"+1{aa['did']}" if len(aa['did']) == 10 else aa['did']
        else:
            update_payload["extension"] = aa['did']

    if schedule_name:
        update_payload["businessSchedule"] = schedule_name

    print(f"  Applying corrections...")
    upd_result = api.call(
        "PUT",
        f"telephony/config/locations/{location_id}/autoAttendants/{aa_id}",
        data=update_payload,
        params={"orgId": api.org_id}
    )

    if "error" in upd_result:
        print(f"  Error applying corrections: {upd_result['error']}")
        return False

    print(f"  Corrections applied successfully.")
    return True


def configure_auto_attendants(api, location_data: dict, filepath: str, read_excel_sheet):
    """
    Main entry point: read the Webex Auto Attendant sheet and create all four AAs.
    """
    from libraries.aso_bulk_import import read_excel_sheet as _read  # noqa: F401

    print(f"\n{'='*60}")
    proceed = input("\nProceed with Auto Attendant configuration? (Y/n): ").strip().lower()
    if proceed not in ['', 'y', 'yes']:
        print("\nAuto Attendant configuration skipped.")
        return

    location_id = location_data['id']

    # Fetch the location's IANA timezone to use as fallback for AAs that
    # don't have an explicit timezone in the sheet
    location_tz = "America/Chicago"  # safe default if API call fails
    loc_detail = api.call("GET", f"locations/{location_id}", params={"orgId": api.org_id})
    if "error" not in loc_detail:
        location_tz = loc_detail.get("timeZone", location_tz)
    print(f"  Location timezone: {location_tz}")

    # -----------------------------------------------------------------------
    # 1. Read sheet
    # -----------------------------------------------------------------------
    sheet_data = read_excel_sheet(filepath, 'Webex Auto Attendant')
    if not sheet_data or len(sheet_data) < 90:
        print("  Error: Could not read 'Webex Auto Attendant' sheet or sheet is too short.")
        return

    # -----------------------------------------------------------------------
    # 2. Parse greeting map and resolve audio files
    # -----------------------------------------------------------------------
    print(f"\n{'='*60}")
    print("Step 1: Resolving announcement audio files")
    print(f"{'='*60}")

    greeting_map = _parse_greeting_map(sheet_data)
    if not greeting_map:
        print("  Warning: No greeting audio file mappings found in sheet (rows 5-8, col M).")

    ann_result = _resolve_announcements(api, location_id, greeting_map)
    if ann_result is None:
        print("\nAuto Attendant configuration aborted due to missing audio files.")
        return
    resolved_announcements, announcement_levels = ann_result

    # -----------------------------------------------------------------------
    # 3. Parse all four AA blocks
    # -----------------------------------------------------------------------
    print(f"\n{'='*60}")
    print("Step 2: Parsing Auto Attendant definitions")
    print(f"{'='*60}")

    aa_list = _parse_aa_blocks(sheet_data, fallback_timezone=location_tz)
    if not aa_list:
        print("  Error: No Auto Attendant blocks found in sheet.")
        return

    print(f"  Found {len(aa_list)} Auto Attendant(s):")
    for aa in aa_list:
        ext_or_did = aa['did'] if aa['did'] else "(no number)"
        print(f"    - {aa['name']}  [{ext_or_did}]  TZ: {aa['timezone']}  Keys: {len(aa['key_configs'])}")

    # -----------------------------------------------------------------------
    # 4. Resolve schedule name — just validate it exists, pass name as string
    # -----------------------------------------------------------------------
    schedule_name = None
    if aa_list and aa_list[0]['schedule_map']:
        first_sched_name = next(iter(aa_list[0]['schedule_map'].values()))
        print(f"\n  Resolving business hours schedule: '{first_sched_name}'...")
        if _validate_schedule_exists(api, location_id, first_sched_name):
            schedule_name = first_sched_name
            print(f"  Schedule confirmed: '{schedule_name}'")
        else:
            print(f"  Warning: Schedule '{first_sched_name}' not found — AA will be created without a schedule.")

    # -----------------------------------------------------------------------
    # 5. Determine greeting block for the main menu
    #    Match the first AA name (or "Main Menu") to the greeting map
    # -----------------------------------------------------------------------
    def _find_greeting_block(aa_name: str) -> dict:
        """
        Find the best matching greeting wav for this AA by matching the menu key
        (col L) against the AA name. Tries full key match first, then last word
        of the key (e.g. 'Main Menu' -> 'main' matches '...Main').
        Falls back to DEFAULT if no match found.
        """
        aa_lower = aa_name.lower()
        for menu_key, wav_filename in greeting_map.items():
            wav_lower = wav_filename.lower()
            if wav_lower not in resolved_announcements:
                continue
            # Full key match (e.g. "directions" in "0387-north canton-directions")
            if menu_key in aa_lower:
                level = announcement_levels.get(wav_lower, "ORGANIZATION")
                return _build_menu_greeting(wav_filename, resolved_announcements[wav_lower], level)
            # Last word of key match (e.g. "main" from "main menu" in "...canton-main")
            last_word = menu_key.split()[-1]
            if last_word and last_word in aa_lower:
                level = announcement_levels.get(wav_lower, "ORGANIZATION")
                return _build_menu_greeting(wav_filename, resolved_announcements[wav_lower], level)
        return {"greeting": "DEFAULT"}

    # -----------------------------------------------------------------------
    # 6. Phase 1 — Create all AA scaffolds (no cross-references yet)
    # -----------------------------------------------------------------------
    print(f"\n{'='*60}")
    print("Step 3: Creating Auto Attendant scaffolds")
    print(f"{'='*60}")

    created_aas = {}   # aa_name -> aa_id

    for aa in aa_list:
        print(f"\n  Auto Attendant: {aa['name']}")

        greeting_block = _find_greeting_block(aa['name'])

        # Check for existing AA with same name
        existing_id = _check_existing_aa(api, location_id, aa['name'])
        if existing_id:
            print(f"  Already exists in Control Hub (ID: {existing_id})")
            _validate_and_update_existing_aa(
                api, location_id, existing_id, aa,
                greeting_block, schedule_name
            )
            created_aas[aa['name']] = existing_id
            continue

        # Does not exist — create it
        if not aa['did']:
            print(f"  Skipping: no phone number or extension defined in sheet — API requires at least one.")
            continue

        payload = _build_aa_payload(aa, greeting_block, schedule_name, location_id)

        print(f"  Creating...")
        result = api.call(
            "POST",
            f"telephony/config/locations/{location_id}/autoAttendants",
            data=payload,
            params={"orgId": api.org_id}
        )

        # Handle phone number already in use (error 4201) — offer extension-only fallback
        if "error" in result:
            raw_error = result.get("error", "")
            is_number_conflict = "4201" in str(raw_error)
            if is_number_conflict and len(aa['did']) >= 10:
                phone_used = f"+1{aa['did']}" if len(aa['did']) == 10 else aa['did']
                print(f"  Error: Phone number {phone_used} is already assigned elsewhere.")
                ext_input = input(
                    f"  Enter an extension to use instead (or press Enter to skip): "
                ).strip()
                if not ext_input:
                    print(f"  Skipped: '{aa['name']}' — no fallback extension provided.")
                    continue
                # Retry with extension only
                fallback_aa = dict(aa, did=ext_input)
                payload = _build_aa_payload(fallback_aa, greeting_block, schedule_name, location_id)
                print(f"  Retrying with extension {ext_input}...")
                result = api.call(
                    "POST",
                    f"telephony/config/locations/{location_id}/autoAttendants",
                    data=payload,
                    params={"orgId": api.org_id}
                )
                if "error" in result:
                    print(f"  Error on retry: {result['error']}")
                    continue
            else:
                print(f"  Error: {raw_error}")
                continue

        aa_id = result.get("id")
        created_aas[aa['name']] = aa_id
        print(f"  Created successfully (ID: {aa_id})")

    if not created_aas:
        print("\n  No Auto Attendants were created.")
        return

    # -----------------------------------------------------------------------
    # 7. Phase 2 — Update each AA to wire cross-AA "Transfer to Sub Menu" keys
    #    Now that all AAs exist, we can resolve their extensions and update menus
    # -----------------------------------------------------------------------
    print(f"\n{'='*60}")
    print("Step 4: Wiring cross-Auto Attendant transfers")
    print(f"{'='*60}")

    # Build extension -> aa_id map from created AAs
    # We need to fetch each AA's details to get its extension
    ext_to_aa_id = {}
    for aa_name, aa_id in created_aas.items():
        if not aa_id:
            continue
        detail = api.call(
            "GET",
            f"telephony/config/locations/{location_id}/autoAttendants/{aa_id}",
            params={"orgId": api.org_id}
        )
        if "error" not in detail:
            ext = str(detail.get("extension", "")).strip()
            phone = _clean_ext(str(detail.get("phoneNumber", "")).strip())
            if ext:
                ext_to_aa_id[ext] = aa_id
            if phone:
                ext_to_aa_id[phone] = aa_id

    updates_needed = False
    for aa in aa_list:
        aa_id = created_aas.get(aa['name'])
        if not aa_id:
            continue

        # Rebuild key configurations, now resolving sub-menu destinations to AA IDs
        key_configurations = []
        for kc in aa['key_configs']:
            action_lower = kc['action'].strip().lower()
            dest_clean   = _clean_ext(kc['dest'])

            if action_lower == "repeat menu":
                key_configurations.append({"key": kc['key'], "action": "REPEAT_MENU"})
                continue

            if action_lower in ("transfer to number", "transfer to sub menu"):
                if not dest_clean:
                    continue
                key_configurations.append({
                    "key":    kc['key'],
                    "action": "TRANSFER_WITHOUT_PROMPT",
                    "value":  dest_clean,
                })
                continue

        if not key_configurations:
            continue

        greeting_block = _find_greeting_block(aa['name'])
        menu = {
            **greeting_block,
            "extensionEnabled":  False,
            "keyConfigurations": key_configurations,
        }

        name_parts = aa['name'].split('-', 2)
        first_name = name_parts[0].strip() if len(name_parts) > 0 else aa['name']
        last_name  = name_parts[-1].strip() if len(name_parts) > 1 else "AA"

        update_payload = {
            "name":              aa['name'],
            "firstName":         first_name,
            "lastName":          last_name,
            "businessHoursMenu": menu,
            "afterHoursMenu":    menu,
        }

        if schedule_name:
            update_payload["businessSchedule"] = schedule_name

        print(f"\n  Updating menu keys for: {aa['name']}")
        upd_result = api.call(
            "PUT",
            f"telephony/config/locations/{location_id}/autoAttendants/{aa_id}",
            data=update_payload,
            params={"orgId": api.org_id}
        )

        if "error" in upd_result:
            print(f"  Warning: Menu update failed - {upd_result['error']}")
        else:
            print(f"  Menu keys updated successfully.")
            updates_needed = True

    # -----------------------------------------------------------------------
    # 8. Summary
    # -----------------------------------------------------------------------
    print(f"\n{'='*60}")
    print("Auto Attendant Configuration Summary")
    print(f"{'='*60}")
    for aa_name, aa_id in created_aas.items():
        status = f"ID: {aa_id}" if aa_id else "FAILED"
        print(f"  {aa_name:<40} {status}")
    print(f"{'='*60}")
