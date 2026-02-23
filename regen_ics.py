#!/usr/bin/env python3
"""Regenerate ICS files from JSON event data (S9, S10)."""

import json
import os
import re

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
        # Find a safe cut point (don't split multi-byte chars)
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

def load_events_from_json(path):
    """Load events from a data/SXX-events.json file."""
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def extract_events_from_html(path):
    """Extract embedded DATA from SXX.html."""
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    # Find the SECOND "var embedded = " (first is notes, second is event data)
    start_marker = 'var embedded = '
    first = content.find(start_marker)
    if first == -1:
        return {}
    idx = content.find(start_marker, first + 1)
    if idx == -1:
        idx = first  # fallback to first if only one
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
    return json.loads(json_str)

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

def generate_ics(name, all_events, all_notes):
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

    for week_num, events in sorted(all_events.items()):
        desc_raw = build_description(all_notes.get(week_num, {}))
        desc_escaped = ics_escape(desc_raw) if desc_raw else ""

        for i, evt in enumerate(events, 1):
            start_str = evt["start"].replace("-", "").replace(":", "").replace("T", "T")
            # Ensure format is YYYYMMDDTHHMMSS
            if len(start_str.split("T")[1]) == 4:
                start_str += "00"
            end_str = evt["end"].replace("-", "").replace(":", "").replace("T", "T")
            if len(end_str.split("T")[1]) == 4:
                end_str += "00"

            lines.append("BEGIN:VEVENT")
            lines.append(f"UID:{s}-s{week_num}-{i}@urban7d")
            lines.append(f"DTSTAMP:{start_str}")
            lines.append(f"DTSTART;TZID=Europe/Paris:{start_str}")
            lines.append(f"DTEND;TZID=Europe/Paris:{end_str}")
            lines.append(fold_line(f"SUMMARY:{evt['label']}"))
            if desc_escaped:
                lines.append(fold_line(f"DESCRIPTION:{desc_escaped}"))
            lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines)


def main():
    # Collect all event data per employee per week
    # Structure: {employee_name: {week_num: [events]}}
    employees = {}
    all_notes = {}

    # Load S9 from JSON
    s9_path = "data/S9-events.json"
    if os.path.exists(s9_path):
        s9_data = load_events_from_json(s9_path)
        all_notes[9] = load_notes(9)
        for name, emp_data in s9_data.items():
            if emp_data.get("events"):
                if name not in employees:
                    employees[name] = {"slug": emp_data.get("slug", slug(name))}
                if "weeks" not in employees[name]:
                    employees[name]["weeks"] = {}
                employees[name]["weeks"][9] = emp_data["events"]

    # Load S10 from embedded HTML
    s10_html = "S10.html"
    if os.path.exists(s10_html):
        s10_data = extract_events_from_html(s10_html)
        all_notes[10] = load_notes(10)
        for name, emp_data in s10_data.items():
            if emp_data.get("events"):
                if name not in employees:
                    employees[name] = {"slug": emp_data.get("slug", slug(name))}
                if "weeks" not in employees[name]:
                    employees[name]["weeks"] = {}
                employees[name]["weeks"][10] = emp_data["events"]

    # Generate ICS files
    os.makedirs(ICS_DIR, exist_ok=True)
    count = 0
    for name, emp_data in sorted(employees.items()):
        s = emp_data.get("slug", slug(name))
        weeks = emp_data.get("weeks", {})
        if not weeks:
            continue

        total_events = sum(len(evts) for evts in weeks.values())
        if total_events == 0:
            continue

        ics_content = generate_ics(name, weeks, all_notes)
        ics_path = os.path.join(ICS_DIR, f"{s}.ics")
        with open(ics_path, 'w', encoding='utf-8', newline='') as f:
            f.write(ics_content)
        count += 1
        print(f"  {s}.ics ({total_events} events)")

    print(f"\n{count} ICS files generated.")


if __name__ == "__main__":
    main()
