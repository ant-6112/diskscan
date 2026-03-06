"""
track_meeting_declines.py
--------------------------
Uses pywin32 to open Outlook, find a specific meeting, and report all declined attendees.

Requirements:
    pip install pywin32

Usage:
    python track_meeting_declines.py
    python track_meeting_declines.py --subject "Weekly Standup"
    python track_meeting_declines.py --subject "Weekly Standup" --date 2026-03-10
"""

import argparse
import sys
from datetime import datetime

try:
    import win32com.client
    import pywintypes
except ImportError:
    print("ERROR: pywin32 is not installed.")
    print("Install it with:  pip install pywin32")
    sys.exit(1)


# ── Outlook constants ────────────────────────────────────────────────────────
OL_MEETING_RESPONSE_DECLINED  = 4
OL_MEETING_RESPONSE_ACCEPTED  = 3
OL_MEETING_RESPONSE_TENTATIVE = 2
OL_MEETING_RESPONSE_NONE      = 5   # No response yet

RESPONSE_LABELS = {
    OL_MEETING_RESPONSE_ACCEPTED:  "Accepted",
    OL_MEETING_RESPONSE_DECLINED:  "Declined",
    OL_MEETING_RESPONSE_TENTATIVE: "Tentative",
    OL_MEETING_RESPONSE_NONE:      "No Response",
}


def connect_to_outlook() -> "win32com.client.CDispatch":
    """Connect to a running (or new) Outlook instance."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("✔  Connected to Outlook.\n")
        return outlook
    except Exception as exc:
        print(f"ERROR: Could not connect to Outlook – {exc}")
        sys.exit(1)


def get_calendar_folder(outlook: "win32com.client.CDispatch") -> "win32com.client.CDispatch":
    """Return the default Calendar folder."""
    namespace = outlook.GetNamespace("MAPI")
    return namespace.GetDefaultFolder(9)   # 9 = olFolderCalendar


def find_meeting(
    calendar,
    subject_keyword: str,
    target_date: datetime | None = None,
) -> "win32com.client.CDispatch | None":
    """
    Search the calendar for a meeting whose subject contains *subject_keyword*.
    If *target_date* is provided, only meetings on that date are considered.
    Returns the first matching AppointmentItem, or None.
    """
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    for item in items:
        try:
            subject: str = item.Subject or ""
            if subject_keyword.lower() not in subject.lower():
                continue

            # Optional date filter
            if target_date:
                item_date = item.Start  # pywintypes.datetime
                if (item_date.year, item_date.month, item_date.day) != (
                    target_date.year,
                    target_date.month,
                    target_date.day,
                ):
                    continue

            return item

        except Exception:
            # Some items (e.g. recurring ghosts) may raise; skip them.
            continue

    return None


def track_declines(meeting) -> None:
    """Print a full attendee response summary with declines highlighted."""
    print("=" * 60)
    print(f"  Meeting : {meeting.Subject}")
    print(f"  Start   : {meeting.Start}")
    print(f"  End     : {meeting.End}")
    print(f"  Location: {meeting.Location or '—'}")
    print(f"  Organizer: {meeting.Organizer}")
    print("=" * 60)

    recipients = meeting.Recipients
    total = recipients.Count

    if total == 0:
        print("No recipients found for this meeting.")
        return

    summary: dict[str, list[str]] = {
        "Declined":    [],
        "Accepted":    [],
        "Tentative":   [],
        "No Response": [],
    }

    for i in range(1, total + 1):          # COM collections are 1-indexed
        try:
            r = recipients.Item(i)
            name   = r.Name
            email  = r.Address if hasattr(r, "Address") else "unknown"
            status = RESPONSE_LABELS.get(r.MeetingResponseStatus, "Unknown")
            summary[status].append(f"{name} <{email}>")
        except Exception as exc:
            print(f"  [warning] Could not read recipient #{i}: {exc}")

    # ── Print declines first (the main focus) ────────────────────────────────
    print(f"\n🔴  DECLINED  ({len(summary['Declined'])})")
    if summary["Declined"]:
        for entry in summary["Declined"]:
            print(f"    • {entry}")
    else:
        print("    (none)")

    # ── Print the rest ───────────────────────────────────────────────────────
    for label in ("Accepted", "Tentative", "No Response"):
        icon = {"Accepted": "🟢", "Tentative": "🟡", "No Response": "⚪"}[label]
        print(f"\n{icon}  {label.upper()}  ({len(summary[label])})")
        for entry in summary[label]:
            print(f"    • {entry}")

    print("\n" + "=" * 60)
    print(f"  Total attendees : {total}")
    print(f"  Declines        : {len(summary['Declined'])}")
    print("=" * 60)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Track declines for an Outlook meeting using pywin32."
    )
    parser.add_argument(
        "--subject",
        default="",
        help="Keyword to search for in the meeting subject (case-insensitive).",
    )
    parser.add_argument(
        "--date",
        default=None,
        help="Optional date filter: YYYY-MM-DD  (e.g. 2026-03-10).",
    )
    args = parser.parse_args()

    # ── If no subject was given, ask interactively ────────────────────────────
    if not args.subject:
        args.subject = input("Enter a keyword from the meeting subject: ").strip()
        if not args.subject:
            print("No subject entered. Exiting.")
            sys.exit(0)

    # ── Parse optional date ───────────────────────────────────────────────────
    target_date: datetime | None = None
    if args.date:
        try:
            target_date = datetime.strptime(args.date, "%Y-%m-%d")
        except ValueError:
            print(f"ERROR: Invalid date format '{args.date}'. Use YYYY-MM-DD.")
            sys.exit(1)

    # ── Main flow ─────────────────────────────────────────────────────────────
    outlook  = connect_to_outlook()
    calendar = get_calendar_folder(outlook)

    print(f"🔍 Searching for meeting containing: '{args.subject}'"
          + (f" on {args.date}" if target_date else "") + " …\n")

    meeting = find_meeting(calendar, args.subject, target_date)

    if meeting is None:
        print("❌  No matching meeting found. Try a different keyword or date.")
        sys.exit(1)

    track_declines(meeting)


if __name__ == "__main__":
    main()
