#!/usr/bin/env python3
"""Microsoft Graph API wrapper for Hermes -- Calendar, Mail, Contacts."""
import json
import sys
import os
from pathlib import Path
from urllib.parse import urlencode

HERMES_HOME = Path.home() / ".hermes"
TOKEN_PATH = HERMES_HOME / "microsoft_token.json"
CLIENT_SECRET_PATH = HERMES_HOME / "microsoft_client_secret.json"
CONFIG_PATH = HERMES_HOME / "microsoft_config.json"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _load_config() -> dict:
    """Load optional user config (sender name, agent name, etc.)."""
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text())
    return {}


def _get_token() -> str:
    """Get a valid access token, refreshing if needed."""
    if not TOKEN_PATH.exists():
        print("ERROR: Not authenticated. Run setup.py --auth-url first.")
        sys.exit(1)

    token = json.loads(TOKEN_PATH.read_text())

    import time
    if token.get("expires_at", 0) < time.time() + 300:
        token = _refresh_token(token)

    return token["access_token"]


def _refresh_token(token: dict) -> dict:
    """Refresh the access token using MSAL."""
    import msal

    config = json.loads(CLIENT_SECRET_PATH.read_text())
    app = msal.PublicClientApplication(
        config["client_id"],
        authority="https://login.microsoftonline.com/consumers",
    )

    result = app.acquire_token_by_refresh_token(
        token["refresh_token"],
        scopes=["Calendars.ReadWrite", "Mail.ReadWrite", "Mail.Send", "User.Read", "Contacts.Read"],
    )

    if "access_token" not in result:
        print(f"ERROR: Token refresh failed: {result.get('error_description', 'Unknown')}")
        sys.exit(1)

    TOKEN_PATH.write_text(json.dumps(result, indent=2))
    return result


def _api_call(method: str, endpoint: str, data: dict = None, params: dict = None):
    """Make an authenticated Graph API call."""
    import urllib.request

    token = _get_token()
    url = f"{GRAPH_BASE}{endpoint}"
    if params:
        url += "?" + urlencode(params)

    req = urllib.request.Request(url, method=method)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type", "application/json")

    if data:
        req.data = json.dumps(data).encode("utf-8")

    try:
        with urllib.request.urlopen(req) as resp:
            if resp.status in (202, 204):
                return {"status": "success"}
            body = resp.read().decode()
            if not body:
                return {"status": "success"}
            return json.loads(body)
    except urllib.error.HTTPError as e:
        error_body = e.read().decode()
        try:
            error_json = json.loads(error_body)
            print(f"ERROR {e.code}: {error_json.get('error', {}).get('message', error_body)}")
        except json.JSONDecodeError:
            print(f"ERROR {e.code}: {error_body}")
        sys.exit(1)


# === CALENDAR ===

def calendar_list(start: str = None, end: str = None, max_results: int = 25, upcoming: bool = True):
    """List calendar events. Defaults to upcoming (future) events only."""
    from datetime import datetime, timedelta

    params = {
        "$top": max_results,
        "$orderby": "start/dateTime",
        "$select": "id,subject,start,end,location,bodyPreview,webLink,isOnlineMeeting,onlineMeeting",
    }

    if not start and upcoming:
        start = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S")
        end = (datetime.utcnow() + timedelta(days=365)).strftime("%Y-%m-%dT%H:%M:%S")

    if start:
        params["startDateTime"] = start
    if end:
        params["endDateTime"] = end
        result = _api_call("GET", "/me/calendarView", params=params)
    else:
        result = _api_call("GET", "/me/events", params=params)

    events = result.get("value", [])
    if not events:
        print("No events found.")
        return []

    output = []
    for e in events:
        start_time = e["start"].get("dateTime", "N/A")
        end_time = e["end"].get("dateTime", "N/A")
        meet_url = e.get("onlineMeeting", {}).get("joinUrl", "") if e.get("onlineMeeting") else ""
        entry = {
            "id": e["id"],
            "subject": e["subject"],
            "start": start_time,
            "end": end_time,
            "notes": e.get("bodyPreview", "")[:80],
            "meet": meet_url,
            "link": e.get("webLink", ""),
        }
        output.append(entry)
        print(f"  {e['subject']}")
        print(f"    When: {start_time} - {end_time}")
        print(f"    ID: {e['id']}")
        if e.get("bodyPreview"):
            print(f"    Notes: {e['bodyPreview'][:80]}")
        if meet_url:
            print(f"    Teams: {meet_url}")
        print()

    return output


def calendar_create(summary: str, start: str, end: str, description: str = "",
                     attendees: list = None, timezone: str = "America/Toronto",
                     add_teams: bool = True):
    """Create a calendar event."""
    event = {
        "subject": summary,
        "body": {"contentType": "HTML", "content": description},
        "start": {"dateTime": start, "timeZone": timezone},
        "end": {"dateTime": end, "timeZone": timezone},
        "isOnlineMeeting": add_teams,
    }

    if attendees:
        event["attendees"] = [
            {"emailAddress": {"address": email}, "type": "required"}
            for email in attendees
        ]

    result = _api_call("POST", "/me/events", data=event)

    meet_url = result.get("onlineMeeting", {}).get("joinUrl", "")
    output = {
        "status": "created",
        "id": result["id"],
        "subject": result["subject"],
        "start": result["start"]["dateTime"],
        "end": result["end"]["dateTime"],
        "link": result.get("webLink", ""),
        "meet": meet_url,
        "attendees": attendees or [],
    }
    print(json.dumps(output, indent=2))
    return output


def calendar_delete(event_id: str):
    """Delete a calendar event."""
    _api_call("DELETE", f"/me/events/{event_id}")
    print(f"Deleted event: {event_id}")


def calendar_invite(summary: str, start: str, end: str, description: str = "",
                    attendees: list = None, timezone: str = "America/Toronto",
                    meet: bool = False):
    """Create a calendar event with attendees and optional video link."""
    from datetime import datetime

    config = _load_config()
    sender_name = config.get("sender_name", "")
    agent_name = config.get("agent_name", "")

    try:
        dt_start = datetime.fromisoformat(start)
        dt_end = datetime.fromisoformat(end)
        duration = int((dt_end - dt_start).total_seconds() / 60)
        date_str = dt_start.strftime("%A, %B %d, %Y")
        time_str = (dt_start.strftime("%I:%M %p").lstrip("0") + " - " +
                    dt_end.strftime("%I:%M %p").lstrip("0") + " (" +
                    timezone.replace("_", " ").split("/")[-1] + ")")
        duration_str = f"{duration} min"
    except Exception:
        date_str = start
        time_str = f"{start} - {end}"
        duration_str = "30 min"

    desc_html = description.replace("\n", "<br>") if description else ""

    # Build sign-off section
    sign_off = ""
    if sender_name:
        sign_off = f'''
        Best regards,<br>
        <span style="font-weight: 500; color: #333;">{sender_name}</span>'''

    agent_footer = ""
    if agent_name and sender_name:
        agent_footer = f'''
      <div style="font-size: 10px; color: #999; letter-spacing: 1.5px; text-transform: uppercase;">Scheduled by</div>
      <div style="font-size: 13px; color: #444; margin-top: 2px;">
        <span style="font-weight: 600;">{agent_name}</span>
        <span style="color: #999;"> &mdash; </span>
        <span style="font-style: italic; color: #666;">{sender_name}&apos;s AI Assistant</span>
      </div>
      <div style="font-size: 10px; color: #aaa; margin-top: 2px;">Powered by Hermes Agent</div>'''

    body_html = f'''<table width="100%" cellpadding="0" cellspacing="0" style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 600px; margin: 0 auto;">
  <tr>
    <td style="padding: 24px; background: #ffffff;">
      <div style="font-size: 13px; color: #888; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 4px;">Topic</div>
      <div style="font-size: 18px; font-weight: 600; color: #1a1a1a; margin-bottom: 16px;">{summary}</div>
      <div style="font-size: 13px; color: #888; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 4px;">Description</div>
      <div style="font-size: 15px; line-height: 1.6; color: #333; margin-bottom: 20px;">
        {desc_html}
      </div>
      <table cellpadding="0" cellspacing="0" style="background: #f5f5f5; border-radius: 8px; padding: 16px; margin-bottom: 20px;">
        <tr>
          <td style="padding: 16px;">
            <table cellpadding="0" cellspacing="0" width="100%">
              <tr>
                <td style="padding-bottom: 8px;">
                  <span style="font-size: 11px; color: #999; text-transform: uppercase; letter-spacing: 1px;">Date</span><br>
                  <span style="font-size: 14px; color: #333; font-weight: 500;">{date_str}</span>
                </td>
              </tr>
              <tr>
                <td style="padding-bottom: 8px;">
                  <span style="font-size: 11px; color: #999; text-transform: uppercase; letter-spacing: 1px;">Time</span><br>
                  <span style="font-size: 14px; color: #333; font-weight: 500;">{time_str}</span>
                </td>
              </tr>
              <tr>
                <td>
                  <span style="font-size: 11px; color: #999; text-transform: uppercase; letter-spacing: 1px;">Duration</span><br>
                  <span style="font-size: 14px; color: #333; font-weight: 500;">{duration_str}</span>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td style="padding: 0 24px;">
      <hr style="border: none; border-top: 1px solid #e5e5e5; margin: 0;">
    </td>
  </tr>{sign_off}
  <tr>
    <td style="padding: 8px 24px 4px 24px; background: #fafafa;">{agent_footer}
    </td>
  </tr>
</table>'''

    event = {
        "subject": summary,
        "body": {"contentType": "HTML", "content": body_html},
        "start": {"dateTime": start, "timeZone": timezone},
        "end": {"dateTime": end, "timeZone": timezone},
        "isOnlineMeeting": meet,
    }

    if attendees:
        event["attendees"] = [
            {"emailAddress": {"address": email}, "type": "required"}
            for email in attendees
        ]

    result = _api_call("POST", "/me/events", data=event)
    meet_url = result.get("onlineMeeting", {}).get("joinUrl", "")

    print(f"Created: {result['subject']}")
    print(f"  When: {result['start']['dateTime']} - {result['end']['dateTime']}")
    print(f"  ID: {result['id']}")
    if meet_url:
        print(f"  Meet: {meet_url}")
    if attendees:
        print(f"  Attendees: {', '.join(attendees)}")
    return result


# === MAIL ===

def mail_list(max_results: int = 10, folder: str = "inbox"):
    """List emails."""
    params = {
        "$top": max_results,
        "$orderby": "receivedDateTime desc",
        "$select": "id,from,subject,receivedDateTime,bodyPreview,isRead",
    }
    result = _api_call("GET", f"/me/mailFolders/{folder}/messages", params=params)
    messages = result.get("value", [])
    if not messages:
        print("No messages found.")
        return
    for m in messages:
        sender = m.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
        read = " " if m.get("isRead") else ">"
        print(f"  {read} {m['subject']}")
        print(f"    From: {sender}")
        print(f"    Date: {m['receivedDateTime']}")
        print(f"    ID: {m['id']}")
        print()


def mail_get(message_id: str):
    """Get full email content."""
    result = _api_call("GET", f"/me/messages/{message_id}")
    print(f"Subject: {result['subject']}")
    print(f"From: {result.get('from', {}).get('emailAddress', {}).get('address', 'N/A')}")
    print(f"Date: {result['receivedDateTime']}")
    print(f"Body:\n{result.get('body', {}).get('content', 'N/A')}")


def mail_send(to: str, subject: str, body: str, html: bool = False, attachment: str = None):
    """Send an email, optionally with a file attachment."""
    import base64
    import mimetypes

    content_type = "HTML" if html else "Text"

    if attachment:
        draft_data = {
            "subject": subject,
            "body": {"contentType": content_type, "content": body},
            "toRecipients": [{"emailAddress": {"address": to}}],
        }
        draft = _api_call("POST", "/me/messages", data=draft_data)
        message_id = draft["id"]

        file_path = os.path.expanduser(attachment)
        if not os.path.exists(file_path):
            print(f"ERROR: File not found: {file_path}")
            sys.exit(1)

        file_size = os.path.getsize(file_path)
        if file_size > 3 * 1024 * 1024:
            print(f"ERROR: Attachment too large ({file_size / 1024 / 1024:.1f} MB). Max 3MB for inline attachments.")
            print("Tip: Compress the image or use a smaller file.")
            sys.exit(1)

        with open(file_path, "rb") as f:
            file_content = base64.b64encode(f.read()).decode("utf-8")

        file_name = os.path.basename(file_path)
        mime_type = mimetypes.guess_type(file_path)[0] or "application/octet-stream"

        attach_data = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": file_name,
            "contentType": mime_type,
            "contentBytes": file_content,
        }
        _api_call("POST", f"/me/messages/{message_id}/attachments", data=attach_data)
        _api_call("POST", f"/me/messages/{message_id}/send", data={})
        print(f"Sent to {to}: {subject} (with attachment: {file_name})")
    else:
        data = {
            "message": {
                "subject": subject,
                "body": {"contentType": content_type, "content": body},
                "toRecipients": [{"emailAddress": {"address": to}}],
            }
        }
        _api_call("POST", "/me/sendMail", data=data)
        print(f"Sent to {to}: {subject}")


# === CONTACTS ===

def contacts_list(max_results: int = 20):
    """List contacts."""
    params = {"$top": max_results, "$select": "displayName,emailAddresses,mobilePhone"}
    result = _api_call("GET", "/me/contacts", params=params)
    contacts = result.get("value", [])
    if not contacts:
        print("No contacts found.")
        return
    for c in contacts:
        emails = ", ".join(e.get("address", "") for e in c.get("emailAddresses", []))
        phone = c.get("mobilePhone", "")
        print(f"  {c['displayName']}")
        if emails:
            print(f"    Email: {emails}")
        if phone:
            print(f"    Phone: {phone}")
        print()


# === CLI ===

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: microsoft_api.py <resource> <action> [args]")
        print("Resources: calendar, mail, contacts")
        print("Actions:")
        print("  calendar list [--start DATETIME] [--end DATETIME] [--max N] [--all]")
        print("  calendar create --summary TITLE --start DATETIME --end DATETIME [--description TEXT] [--attendees EMAIL1,EMAIL2]")
        print("  calendar invite --summary TITLE --start DATETIME --end DATETIME [--description TEXT] [--attendees EMAIL1,EMAIL2] [--meet]")
        print("  calendar delete EVENT_ID")
        print("  mail list [--max N]")
        print("  mail get MESSAGE_ID")
        print("  mail send --to EMAIL --subject TEXT --body TEXT [--html] [--attachment FILEPATH]")
        print("  contacts list [--max N]")
        sys.exit(1)

    resource = sys.argv[1]
    action = sys.argv[2] if len(sys.argv) > 2 else "list"

    if resource == "calendar":
        if action == "list":
            start = None
            end = None
            max_r = 25
            upcoming = True
            i = 3
            while i < len(sys.argv):
                if sys.argv[i] == "--start" and i + 1 < len(sys.argv):
                    start = sys.argv[i + 1]; i += 2
                elif sys.argv[i] == "--end" and i + 1 < len(sys.argv):
                    end = sys.argv[i + 1]; i += 2
                elif sys.argv[i] == "--max" and i + 1 < len(sys.argv):
                    max_r = int(sys.argv[i + 1]); i += 2
                elif sys.argv[i] == "--all":
                    upcoming = False; i += 1
                else:
                    i += 1
            calendar_list(start, end, max_r, upcoming)

        elif action == "create":
            summary = desc = start = end = ""
            attendees = None
            tz = "America/Toronto"
            i = 3
            while i < len(sys.argv):
                if sys.argv[i] == "--summary" and i + 1 < len(sys.argv):
                    summary = sys.argv[i + 1]; i += 2
                elif sys.argv[i] == "--description" and i + 1 < len(sys.argv):
                    desc = sys.argv[i + 1]; i += 2
                elif sys.argv[i] == "--start" and i + 1 < len(sys.argv):
                    start = sys.argv[i + 1]; i += 2
                elif sys.argv[i] == "--end" and i + 1 < len(sys.argv):
                    end = sys.argv[i + 1]; i += 2
                elif sys.argv[i] == "--attendees" and i + 1 < len(sys.argv):
                    attendees = sys.argv[i + 1].split(","); i += 2
                else:
                    i += 1
            if not all([summary, start, end]):
                print("ERROR: --summary, --start, and --end are required")
                sys.exit(1)
            calendar_create(summary, start, end, desc, attendees, tz)

        elif action == "delete":
            if len(sys.argv) < 4:
                print("Usage: calendar delete EVENT_ID")
                sys.exit(1)
            calendar_delete(sys.argv[3])

        elif action == "invite":
            summary = desc = start = end = ""
            attendees = None
            tz = "America/Toronto"
            use_meet = False
            i = 3
            while i < len(sys.argv):
                if sys.argv[i] == "--summary" and i + 1 < len(sys.argv):
                    summary = sys.argv[i + 1]; i += 2
                elif sys.argv[i] == "--description" and i + 1 < len(sys.argv):
                    desc = sys.argv[i + 1]; i += 2
                elif sys.argv[i] == "--start" and i + 1 < len(sys.argv):
                    start = sys.argv[i + 1]; i += 2
                elif sys.argv[i] == "--end" and i + 1 < len(sys.argv):
                    end = sys.argv[i + 1]; i += 2
                elif sys.argv[i] == "--attendees" and i + 1 < len(sys.argv):
                    attendees = sys.argv[i + 1].split(","); i += 2
                elif sys.argv[i] == "--meet":
                    use_meet = True; i += 1
                else:
                    i += 1
            if not all([summary, start, end]):
                print("ERROR: --summary, --start, and --end are required")
                sys.exit(1)
            calendar_invite(summary, start, end, desc, attendees, tz, use_meet)

    elif resource == "mail":
        if action == "list":
            max_r = 10
            i = 3
            while i < len(sys.argv):
                if sys.argv[i] == "--max" and i + 1 < len(sys.argv):
                    max_r = int(sys.argv[i + 1]); i += 2
                else:
                    i += 1
            mail_list(max_r)
        elif action == "get":
            if len(sys.argv) < 4:
                print("Usage: mail get MESSAGE_ID")
                sys.exit(1)
            mail_get(sys.argv[3])
        elif action == "send":
            to = subject = body = ""
            html = False
            attachment = None
            i = 3
            while i < len(sys.argv):
                if sys.argv[i] == "--to" and i + 1 < len(sys.argv):
                    to = sys.argv[i + 1]; i += 2
                elif sys.argv[i] == "--subject" and i + 1 < len(sys.argv):
                    subject = sys.argv[i + 1]; i += 2
                elif sys.argv[i] == "--body" and i + 1 < len(sys.argv):
                    body = sys.argv[i + 1]; i += 2
                elif sys.argv[i] == "--html":
                    html = True; i += 1
                elif sys.argv[i] == "--attachment" and i + 1 < len(sys.argv):
                    attachment = sys.argv[i + 1]; i += 2
                else:
                    i += 1
            if not all([to, subject, body]):
                print("ERROR: --to, --subject, and --body are required")
                sys.exit(1)
            mail_send(to, subject, body, html, attachment)

    elif resource == "contacts":
        max_r = 20
        i = 3
        while i < len(sys.argv):
            if sys.argv[i] == "--max" and i + 1 < len(sys.argv):
                max_r = int(sys.argv[i + 1]); i += 2
            else:
                i += 1
        contacts_list(max_r)

    else:
        print(f"Unknown resource: {resource}")
        sys.exit(1)
