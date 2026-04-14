# Microsoft Workspace Feature Plan

## Positioning

**Core identity:** Email skill for Hermes Agent
**Secondary:** Calendar, contacts, and productivity integrations

The skill should be known as "the email skill" first. Everything else is a bonus that makes it more useful over time.

## Build Order

### Phase 1: Email Polish (ship this first)
These make the email experience complete. Do these before anything else.

| Feature | Endpoint | Priority | Effort |
|---------|----------|----------|--------|
| Mail search | `/me/messages?$search="term"` | HIGH | Small |
| Mail filter (unread, important) | `/me/messages?$filter=...` | HIGH | Small |
| Reply to email | `/me/messages/{id}/reply` | HIGH | Small |
| Forward email | `/me/messages/{id}/forward` | HIGH | Small |
| Mail folders list | `/me/mailFolders` | MED | Small |
| Move email to folder | `/me/messages/{id}/move` | MED | Small |

**CLI additions:**
```
mail search --query "keyword" [--max N]
mail reply --id MSG_ID --body "text"
mail forward --id MSG_ID --to "email" [--body "text"]
mail folders
mail move --id MSG_ID --folder "folder_name"
```

**Estimated time:** 2-3 hours
**Blocks main goal:** No — these directly improve email

### Phase 2: Calendar Upgrades (ship after Phase 1)

| Feature | Endpoint | Priority | Effort |
|---------|----------|----------|--------|
| Free/busy check | `/me/calendar/getSchedule` | MED | Medium |
| Calendar availability | Find open slots | MED | Medium |
| Update event | PATCH `/me/events/{id}` | MED | Small |

**CLI additions:**
```
calendar freebusy --start DATE --end DATE [--attendees email1,email2]
calendar slots --date DATE [--duration 30]
calendar update --id EVENT_ID [--summary NEW] [--start NEW]
```

**Estimated time:** 2-3 hours
**Blocks main goal:** No

### Phase 3: Bonus Features (ship when ready, no rush)

#### 3a. To Do / Tasks
```
todo lists                          # List task lists
todo tasks --list LIST_ID           # List tasks in a list
todo add --list LIST_ID --title "Task"
todo complete --task TASK_ID
```

**API:** `/me/todo/lists`, `/me/todo/lists/{id}/tasks`
**New permission needed:** `Tasks.ReadWrite`
**Estimated time:** 2 hours

#### 3b. User Profile
```
profile                             # Show user info
```

**API:** `/me`
**No new permissions needed**
**Estimated time:** 30 min

#### 3c. OneDrive / Files
```
files list [--folder PATH]
files search --query "term"
files download --id FILE_ID --output PATH
files upload --file PATH [--folder PATH]
```

**API:** `/me/drive/root/children`, `/me/drive/search`
**New permission needed:** `Files.ReadWrite`
**Estimated time:** 3-4 hours

#### 3d. Free/Busy (if not done in Phase 2)
Same as Phase 2, moved here if deprioritized.

## Permission Summary

Current permissions (already configured):
- Calendars.ReadWrite
- Mail.ReadWrite
- Mail.Send
- User.Read
- Contacts.Read

New permissions needed for Phase 3:
- Tasks.ReadWrite (To Do)
- Files.ReadWrite (OneDrive)

No new permissions needed for Phase 1 or 2.

## Design Principles

1. **Email first** — every feature should tie back to email workflows
2. **Small, testable increments** — each phase ships independently
3. **No personal info** — keep the skill generic and configurable
4. **CLI consistency** — follow the existing `resource action --flag` pattern
5. **Safe by default** — preview before sending, validate before destructive ops

## Release Strategy

- Phase 1: Tag v1.1.0 — "Complete email suite"
- Phase 2: Tag v1.2.0 — "Calendar intelligence"
- Phase 3a: Tag v1.3.0 — "Task management"
- Phase 3b: Tag v1.3.1 — "User profile"
- Phase 3c: Tag v1.4.0 — "File management"

Each release is a git tag + updated SKILL.md + tested end-to-end.

## What NOT to build

- Teams integration (different auth, complex, low value for personal use)
- SharePoint (enterprise-focused, not personal assistant relevant)
- Excel manipulation (too niche, use OneDrive download instead)
- Planner (group-focused, To Do is better for personal)
