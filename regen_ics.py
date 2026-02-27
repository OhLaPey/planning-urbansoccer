#!/usr/bin/env python3
"""Regenerate ICS files from HTML event data (all weeks, auto-discovered)."""

import glob
import json
import os
import re
from datetime import datetime, timezone


ICS_DIR = "ics"
NOTES_DIR = "notes"


def slug(name):
    s = name.lower()
    for old, new in [("ï", "i"), ("é", "e"), ("è", "e"), ("ê", "e"),
                     ("ô", "o"), ("ü", "u"), ("ù", "u"), ("û", "u"),
                     ("à", "a"), ("â", "a"), ("ç", "c")]:
        s = s.replace(old, new)
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")


def ics_escape(text):
    return text.replace("\\", "\\\\").replace("\n", "\\n").replace(",", "\\,").replace(";", "\\;")


def fold_line(line, max_len=75):
    """Fold long lines per RFC 5545 (max 75 octets per line)."""
    encoded = line.encode('utf-8')
    if len(encoded) <= max_len:
        return line
    parts = []
    while len(encoded) > max_len:
        cut = max_len if not parts else max_len - 1  # -1 for leading space
        while cut > 0 and (encoded[cut] & 0xC0) == 0x80:
            cut -= 1
        parts.append(encoded[:cut].decode('utf-8'))
        encoded = encoded[cut:]
    if encoded:
        parts.append(encoded.decode('utf-8'))
    return ("\r\n ").join(parts)


def load_notes(week_num):
    path = os.path.join(NOTES_DIR, f"S{week_num}.json")
    if not os.path.exists(path):
        return {}
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def extract_events_from_html(path):
    """Extract embedded event DATA from SXX.html."""
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    # Find "var embedded = " that contains employee event data (has "slug" and "events" keys)
    start_marker = 'var embedded = '
    pos = 0
    while True:
        idx = content.find(start_marker, pos)
        if idx == -1:
            return {}
        idx += len(start_marker)
        # Find matching closing brace
        depth = 0
        end_idx = idx
        for i in range(idx, len(content)):
            if content[i] == '{':
                depth += 1
            elif content[i] == '}':
                depth -= 1
                if depth == 0:
                    end_idx = i + 1
                    break
        json_str = content[idx:end_idx]
        try:
            data = json.loads(json_str)
            # Check if this looks like event data (has employee names with "slug" and "events")
            first_val = next(iter(data.values()), None)
            if isinstance(first_val, dict) and "events" in first_val:
                return data
        except (json.JSONDecodeError, StopIteration):
            pass
        pos = idx


def build_description(notes):
    """Build description text from notes."""
    if not notes:
        return ""
    desc = notes.get("comment", "")
    for upd in notes.get("updates", []):
        upd_text = upd.get("text", "")
        upd_date = upd.get("date", "")
        if upd_text:
            prefix = f"MAJ {upd_date}: " if upd_date else "MAJ: "
            if desc:
                desc += "\n"
            desc += prefix + upd_text
    return desc


def generate_ics(name, all_events, all_notes, dtstamp_utc):
    """Generate ICS content for one employee across all weeks."""
    s = slug(name)
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Planning Urban 7D//FR",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:Planning {name}",
        "X-WR-TIMEZONE:Europe/Paris",
        "REFRESH-INTERVAL;VALUE=DURATION:PT12H",
        "X-PUBLISHED-TTL:PT12H",
        "BEGIN:VTIMEZONE",
        "TZID:Europe/Paris",
        "BEGIN:STANDARD",
        "TZOFFSETFROM:+0200",
        "TZOFFSETTO:+0100",
        "TZNAME:CET",
        "DTSTART:19701025T030000",
        "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU",
        "END:STANDARD",
        "BEGIN:DAYLIGHT",
        "TZOFFSETFROM:+0100",
        "TZOFFSETTO:+0200",
        "TZNAME:CEST",
        "DTSTART:19700329T020000",
        "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU",
        "END:DAYLIGHT",
        "END:VTIMEZONE",
    ]

    for week_num in sorted(all_events.keys()):
        events = all_events[week_num]
        desc_raw = build_description(all_notes.get(week_num, {}))
        desc_escaped = ics_escape(desc_raw) if desc_raw else ""

        for i, evt in enumerate(events, 1):
            start_str = evt["start"].replace("-", "").replace(":", "")
            end_str = evt["end"].replace("-", "").replace(":", "")
            # Ensure format is YYYYMMDDTHHMMSS
            if len(start_str.split("T")[1]) == 4:
                start_str += "00"
            if len(end_str.split("T")[1]) == 4:
                end_str += "00"

            summary = evt['label'].replace("\\", "\\\\").replace(",", "\\,").replace(";", "\\;")

            lines.append("BEGIN:VEVENT")
            lines.append(f"UID:{s}-s{week_num}-{i}@urban7d")
            lines.append(f"DTSTAMP:{dtstamp_utc}")
            lines.append(f"DTSTART;TZID=Europe/Paris:{start_str}")
            lines.append(f"DTEND;TZID=Europe/Paris:{end_str}")
            lines.append(fold_line(f"SUMMARY:{summary}"))
            if desc_escaped:
                lines.append(fold_line(f"DESCRIPTION:{desc_escaped}"))
            lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"


def main():
    # DTSTAMP must be UTC per RFC 5545
    dtstamp_utc = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    # Auto-discover all week HTML files (S3.html, S5.html, S9.html, S10.html, ...)
    html_files = sorted(glob.glob("S*.html"))
    print(f"Found {len(html_files)} week files: {', '.join(html_files)}")

    # Collect all event data per employee per week
    # Structure: {employee_name: {week_num: [events]}}
    employees = {}
    all_notes = {}

    for html_file in html_files:
        # Extract week number from filename (S9.html -> 9, S10.html -> 10)
        match = re.match(r"S(\d+)\.html", html_file)
        if not match:
            continue
        week_num = int(match.group(1))

        data = extract_events_from_html(html_file)
        if not data:
            print(f"  {html_file}: no event data found, skipping")
            continue

        all_notes[week_num] = load_notes(week_num)
        emp_count = 0

        for name, emp_data in data.items():
            events = emp_data.get("events", [])
            if not events:
                continue
            if name not in employees:
                employees[name] = {"slug": emp_data.get("slug", slug(name)), "weeks": {}}
            employees[name]["weeks"][week_num] = events
            emp_count += 1

        print(f"  {html_file}: week {week_num}, {emp_count} employees with events")

    # Generate ICS files
    os.makedirs(ICS_DIR, exist_ok=True)
    count = 0
    for name in sorted(employees.keys()):
        emp_data = employees[name]
        s = emp_data["slug"]
        weeks = emp_data.get("weeks", {})
        if not weeks:
            continue

        total_events = sum(len(evts) for evts in weeks.values())
        if total_events == 0:
            continue

        ics_content = generate_ics(name, weeks, all_notes, dtstamp_utc)
        ics_path = os.path.join(ICS_DIR, f"{s}.ics")
        with open(ics_path, 'w', encoding='utf-8', newline='') as f:
            f.write(ics_content)
        count += 1
        week_list = ",".join(f"S{w}" for w in sorted(weeks.keys()))
        print(f"  {s}.ics ({total_events} events, weeks: {week_list})")

    print(f"\n{count} ICS files generated.")


if __name__ == "__main__":
    main()
