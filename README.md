# Folder Monitor

A lightweight Windows desktop app that monitors one or more folders and automatically deletes files when they exceed a configured size limit. Includes a tray icon, startup integration, and a built-in locked file inspector.

---

## Overview

Folder Monitor keeps selected folders under a defined storage threshold.  
When a folder grows beyond its limit, the app deletes files based on modified time (oldest or newest first) until the folder is back under the limit.

It is designed for:

- Cache folders  
- Recording/output folders  
- Temp storage directories  
- Any location that must stay below a size cap  

---

## Features

- Monitor **multiple folders**
- Custom size limit per folder (`b`, `kb`, `mb`, `gb`, `tb`)
- Default unit is **GB** if none is provided
- Delete strategy:
  - **Earliest (oldest) first**
  - **Latest (newest) first**
- Automatic periodic checks (default: every 5 seconds)
- Manual “Run Check Now” button
- System tray support
- Optional “Start on Windows startup”
- Built-in **Locked File Inspector**
  - Detects processes locking a file
  - Displays PID, process name, type, and service
  - Option to terminate locking processes
  - Retry deletion after unlock
- Live activity log
- Config file persistence
- Automatic migration of legacy configs

---

## How It Works

1. The app scans each monitored folder.
2. It calculates total folder size recursively.
3. If the size exceeds the configured limit:
   - Files are sorted by modified time.
   - Files are deleted until the folder is back under limit.
4. If a file is locked:
   - A notification appears.
   - You can open the Locked File Inspector to view and manage locking processes.

---

## Size Input Examples

| Input  | Meaning      |
|--------|--------------|
| `0.5`  | 0.5 GB       |
| `750mb`| 750 MB       |
| `120kb`| 120 KB       |
| `1.2gb`| 1.2 GB       |
| `2tb`  | 2 TB         |

If you do not specify a unit, **GB is assumed**.

**DESIGNED FOR WINDOWS. MAY BE BUGGY ON OTHER OPERATING SYSTEMS**
