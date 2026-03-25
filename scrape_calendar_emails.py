#!/usr/bin/env python3
"""
Calendar Event Email Scraper
Extracts unique attendee/guest emails from a Google Calendar .ics export.

Usage:
    python scrape_calendar_emails.py <path_to_ics_file> [--output emails.csv] [--exclude-own YOUR_EMAIL]
"""

import argparse
import csv
import re
import sys
from pathlib import Path
from collections import defaultdict


def parse_ics_attendees(ics_path: str) -> list[dict]:
    """
    Parse an .ics file and extract attendee information.
    Returns a list of dicts with email, name, and event details.
    """
    content = Path(ics_path).read_text(encoding="utf-8")

    # Unfold long lines (RFC 5545: lines starting with space/tab are continuations)
    content = re.sub(r"\r?\n[ \t]", "", content)

    events = []
    current_event = None

    for line in content.splitlines():
        if line.strip() == "BEGIN:VEVENT":
            current_event = {"summary": "", "attendees": [], "start": ""}
        elif line.strip() == "END:VEVENT":
            if current_event:
                events.append(current_event)
            current_event = None
        elif current_event is not None:
            if line.startswith("SUMMARY"):
                current_event["summary"] = line.split(":", 1)[-1].strip()
            elif line.startswith("DTSTART"):
                current_event["start"] = line.split(":", 1)[-1].strip()
            elif "ATTENDEE" in line.upper():
                attendee = _parse_attendee_line(line)
                if attendee:
                    current_event["attendees"].append(attendee)
            elif line.startswith("ORGANIZER"):
                organizer = _parse_attendee_line(line)
                if organizer:
                    organizer["role"] = "ORGANIZER"
                    current_event["attendees"].append(organizer)

    return events


def _parse_attendee_line(line: str) -> dict | None:
    """Extract email and common name from an ATTENDEE or ORGANIZER line."""
    # Email is in the mailto: portion
    email_match = re.search(r"mailto:([^\s;\"]+)", line, re.IGNORECASE)
    if not email_match:
        return None

    email = email_match.group(1).strip().lower()

    # Try to extract CN (common name) parameter
    cn_match = re.search(r'CN=([^;:]+)', line, re.IGNORECASE)
    name = cn_match.group(1).strip().strip('"') if cn_match else ""

    # Role
    role_match = re.search(r'ROLE=([^;:]+)', line, re.IGNORECASE)
    role = role_match.group(1).strip() if role_match else ""

    return {"email": email, "name": name, "role": role}


def build_email_list(events: list[dict], exclude_emails: set[str] = None) -> list[dict]:
    """
    Build a deduplicated email list from parsed events.
    Tracks how many events each contact appeared in and their most recent event.
    """
    exclude_emails = {e.lower() for e in (exclude_emails or set())}
    contacts = defaultdict(lambda: {"name": "", "events_count": 0, "last_event": "", "last_date": ""})

    for event in events:
        for att in event["attendees"]:
            email = att["email"]
            if email in exclude_emails:
                continue
            # Skip resource/room calendars
            if "resource.calendar.google.com" in email:
                continue

            entry = contacts[email]
            entry["events_count"] += 1
            if att["name"] and not entry["name"]:
                entry["name"] = att["name"]
            # Track most recent event
            if event["start"] >= entry["last_date"]:
                entry["last_date"] = event["start"]
                entry["last_event"] = event["summary"]

    # Sort by frequency (most events first), then alphabetically
    result = []
    for email, info in sorted(contacts.items(), key=lambda x: (-x[1]["events_count"], x[0])):
        result.append({
            "email": email,
            "name": info["name"],
            "events_count": info["events_count"],
            "last_event": info["last_event"],
        })

    return result


def write_csv(contacts: list[dict], output_path: str):
    """Write contacts to a CSV file."""
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["email", "name", "events_count", "last_event"])
        writer.writeheader()
        writer.writerows(contacts)


def write_txt(contacts: list[dict], output_path: str):
    """Write just the emails, one per line (easy to paste into email tools)."""
    with open(output_path, "w", encoding="utf-8") as f:
        for c in contacts:
            f.write(c["email"] + "\n")


def main():
    parser = argparse.ArgumentParser(
        description="Extract attendee emails from a Google Calendar .ics export."
    )
    parser.add_argument("ics_file", help="Path to the .ics file")
    parser.add_argument(
        "-o", "--output", default="client_emails.csv",
        help="Output file path (default: client_emails.csv). Use .txt for plain email list."
    )
    parser.add_argument(
        "-e", "--exclude", nargs="*", default=[],
        help="Email addresses to exclude (e.g., your own email)"
    )
    parser.add_argument(
        "--emails-only", action="store_true",
        help="Output only emails, one per line (forces .txt format)"
    )

    args = parser.parse_args()

    if not Path(args.ics_file).exists():
        print(f"Error: File not found: {args.ics_file}", file=sys.stderr)
        sys.exit(1)

    # Parse
    print(f"Parsing {args.ics_file}...")
    events = parse_ics_attendees(args.ics_file)
    print(f"  Found {len(events)} events")

    # Build deduplicated list
    contacts = build_email_list(events, exclude_emails=set(args.exclude))
    print(f"  Found {len(contacts)} unique email addresses")

    if not contacts:
        print("No attendee emails found. Check that your .ics export includes guest info.")
        sys.exit(0)

    # Write output
    if args.emails_only or args.output.endswith(".txt"):
        output_path = args.output if args.output.endswith(".txt") else args.output.rsplit(".", 1)[0] + ".txt"
        write_txt(contacts, output_path)
    else:
        output_path = args.output
        write_csv(contacts, output_path)

    print(f"  Saved to {output_path}")

    # Preview
    print(f"\nTop contacts:")
    for c in contacts[:10]:
        name_part = f" ({c['name']})" if c["name"] else ""
        print(f"  {c['email']}{name_part} — {c['events_count']} event(s)")

    if len(contacts) > 10:
        print(f"  ... and {len(contacts) - 10} more")


if __name__ == "__main__":
    main()
