# Microsoft Workspace Skill for Hermes Agent

```
                    *                  *                       *
             *           _____________           * .'
                        |\           /|
                        | \    @    / |
                        |  \  ___  /  |
                        |   \/   \/   |
                        |    \___/    |
                        |     | |     |
                        |_____| |_____|

                 __  ____________  __  ______________
                / / / / ____/ __ \/  |/  / ____/ ___/
               / /_/ / __/ / /_/ / /|_/ / __/  \__ \
              / __  / /___/ _, _/ /  / / /___ ___/ /
             /_/ /_/_____/_/ |_/_/  /_/_____//____/

            ┌─────────────────────────────────────┐
            │  Outlook · Calendar · Email · Tasks │
            └─────────────────────────────────────┘
```

Outlook Calendar, Email, Contacts, and To-Do integration via Microsoft Graph API. Works with Hotmail, Outlook.com, and Microsoft 365 accounts.

## Features

### Email (primary focus)
- List emails with filters (unread, important, by folder)
- Search emails by keyword
- Read full email content
- Send emails with file attachments (up to 3MB)
- Reply and reply-all to emails
- Forward emails
- List and navigate mail folders
- Move emails between folders
- Safe email sender with preview mode (prevents shell variable issues)

### Calendar
- List, create, update, delete calendar events
- Send calendar invites with branded HTML templates and Teams video links
- Check free/busy status for one or more people
- Find open time slots where all attendees are free

### Contacts & Profile
- List contacts
- Get user profile info

## Prerequisites

- Python 3.8+
- `msal` package (`pip install msal`)
- A Microsoft account (Hotmail, Outlook.com, or Microsoft 365)
- Azure App Registration (free, takes 5 minutes)

## Setup

### Step 1: Create an Azure App Registration

1. Go to [https://portal.azure.com](https://portal.azure.com)
2. Sign in with your Microsoft account
3. Navigate to **Azure Active Directory** > **App registrations** > **New registration**
4. Fill in:
   - **Name:** `Hermes Agent` (or any name you like)
   - **Supported account types:** Select **"Accounts in any organizational directory and personal Microsoft accounts (for Hotmail/Outlook.com)"**
   - **Redirect URI:** Select **"Public client/native (mobile & desktop)"** and enter `http://localhost`
5. Click **Register**
6. Copy these values from the overview page:
   - **Application (client) ID** -- you will need this
   - **Directory (tenant) ID** -- you will need this
7. Go to **Certificates & secrets** > **New client secret** > Add > Copy the **Value** (not the ID)
8. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**
9. Add these permissions:
   - `Calendars.ReadWrite`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `User.Read`
   - `Contacts.Read`
10. Click **Grant admin consent** (if available)

### Step 2: Save your credentials

Use the auth script to save your credentials:

```bash
./scripts/auth.sh --client-id "YOUR_CLIENT_ID" --client-secret "YOUR_CLIENT_SECRET" --tenant-id "YOUR_TENANT_ID"
```

Or if using a personal Microsoft account (Hotmail/Outlook.com), you can omit the tenant ID:

```bash
./scripts/auth.sh --client-id "YOUR_CLIENT_ID" --client-secret "YOUR_CLIENT_SECRET"
```

Credentials are saved to `~/.hermes/microsoft_client_secret.json`.

### Step 3: Authorize your account

```bash
# Get the auth URL
python3 scripts/setup.py --auth-url

# Open the URL in your browser, sign in, and copy the redirect URL

# Paste the redirect URL
python3 scripts/setup.py --auth-code 'PASTE_THE_REDIRECT_URL_HERE'
```

The token is saved to `~/.hermes/microsoft_token.json` and auto-refreshes.

### Step 4: (Optional) Set your display name

If you want personalized calendar invite templates:

```bash
./scripts/auth.sh --config --name "Jane Doe" --agent "Hermes"
```

This saves to `~/.hermes/microsoft_config.json`. Without this, calendar invites use a clean template with no sign-off.

### Step 5: Verify

```bash
./scripts/auth.sh --check
```

## Usage

### Email

```bash
API="python3 scripts/microsoft_api.py"

# List recent emails
$API mail list --max 10

# List unread emails only
$API mail list --unread

# List high-importance emails
$API mail list --important

# List emails from a specific folder
$API mail list --folder archive --max 5

# Search emails by keyword
$API mail search --query "invoice" --max 5
$API mail search --query "project report" --folder sent

# Read an email
$API mail get MESSAGE_ID

# Send an email
$API mail send --to "user@example.com" --subject "Hello" --body "How are you?"

# Send with attachment (max 3MB)
$API mail send --to "user@example.com" --subject "See attached" --body "Here's the file" --attachment /path/to/file.png

# Reply to an email
$API mail reply MESSAGE_ID --body "Thanks for the update!"

# Reply-all to an email
$API mail replyall MESSAGE_ID --body "Sounds good, team."

# Forward an email
$API mail forward MESSAGE_ID --to "forward@example.com" --body "FYI"

# List all mail folders
$API mail folders

# Move an email to a different folder
$API mail move MESSAGE_ID --folder FOLDER_ID
```

### Safe Email Sending (recommended)

The safe sender previews emails before sending. This prevents issues with shell `$` variable interpretation:

```bash
SAFE="scripts/safe_mail_send.sh"

# Preview only (does not send)
$SAFE --to "user@example.com" --subject "Invoice for $20" --body-file /tmp/email.txt

# Preview and send
$SAFE --to "user@example.com" --subject "Invoice for $20" --body-file /tmp/email.txt --confirm

# With attachment
$SAFE --to "user@example.com" --subject "Report" --body-file /tmp/email.txt --attachment /path/to/file.png --confirm
```

### Calendar

```bash
# List upcoming events (default)
$API calendar list

# List all events including past
$API calendar list --all

# List events in a date range
$API calendar list --start "2026-04-10T00:00:00" --end "2026-04-10T23:59:59"

# Create an event with Teams link
$API calendar create --summary "Meeting" --start "2026-04-10T15:00:00-04:00" --end "2026-04-10T15:30:00-04:00" --description "Sync up"

# Create an invite with attendees
$API calendar invite --summary "Project Sync" --start "2026-04-11T14:00:00-04:00" --end "2026-04-11T14:30:00-04:00" --description "Let's discuss" --attendees "john@example.com,jane@example.com" --meet

# Update an event (partial -- only update what you pass)
$API calendar update EVENT_ID --summary "New Title"
$API calendar update EVENT_ID --start "2026-04-11T16:00:00-04:00" --end "2026-04-11T16:30:00-04:00"

# Delete an event
$API calendar delete EVENT_ID

# Check free/busy for people (30-min blocks)
$API calendar freebusy --emails "user@example.com,other@example.com" --start "2026-04-14T09:00:00-04:00" --end "2026-04-14T17:00:00-04:00" --interval 30

# Find open 30-min slots where all attendees are free (15-min blocks)
$API calendar findopen --emails "user@example.com" --start "2026-04-14T09:00:00-04:00" --end "2026-04-14T17:00:00-04:00" --duration 30 --interval 15
```

### Contacts & Profile

```bash
# List contacts
$API contacts list --max 20

# Get your profile info
$API user profile
```

## File Locations

| File | Location | Purpose |
|------|----------|---------|
| Credentials | `~/.hermes/microsoft_client_secret.json` | Azure app credentials |
| Token | `~/.hermes/microsoft_token.json` | OAuth2 token (auto-refreshed) |
| Config | `~/.hermes/microsoft_config.json` | Display name, agent name (optional) |

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `Not authenticated` | Run `scripts/setup.py --auth-url` then `--auth-code` |
| `Token refresh failed` | Re-authorize: run `scripts/setup.py --auth-url` |
| `AADSTS700016` | Wrong client_id or tenant_id in credentials |
| `AADSTS65001` | Missing API permissions -- check Azure portal |
| `$20` becomes `0` in subject | Use `safe_mail_send.sh` or write body to a temp file with `cat << 'EOF'` |
| Attachment too large | Max 3MB for inline attachments. Compress or split the file |
| `Insufficient privileges` | Need `Calendars.ReadWrite` permission in Azure |

## Hermes Integration

To install this skill in Hermes Agent:

```bash
# Copy to skills directory
cp -r . ~/.hermes/skills/productivity/microsoft-workspace

# Or install from a published repo
hermes skills install your-username/microsoft-workspace-skill
```

Once installed, the agent can use it to send emails, manage calendar events, and look up contacts on your behalf.

## License

MIT
