# EntryPass-MiHCM Sync

Portable attendance data sync tool — bridges EntryPass access control systems with MiHCM cloud HR platform.

**By Dajayana Trading** | www.dajayana.com

## Features

- Automated EntryPass → MiHCM attendance data sync
- OAuth 2.0 API authentication with auto token refresh
- Batch upload with retry logic and exponential backoff
- Duplicate detection — safe to re-run
- Multi-site support with location tagging
- Windows Task Scheduler integration
- GUI setup wizard (PowerShell WinForms)
- License-controlled distribution
- Complete audit logging

## Quick Start

1. Copy the `app/` folder to the target PC
2. Right-click `setup.ps1` → Run with PowerShell
3. Enter your license key, API credentials, and site details
4. Click Install — Task Scheduler is configured automatically

## Structure

```
app/
├── setup.ps1          — GUI setup wizard
├── sync.ps1           — Main sync engine
├── config.json        — Auto-generated configuration
├── run.bat            — Manual run launcher
└── uninstall.ps1      — Remove scheduled task
```

## Licensing

This software requires a valid license key issued by Dajayana Trading.
Contact: rgyong@outlook.com | +60 16-883 8338

## Documentation

- `docs/Solution_Overview.pdf` — Technical solution overview
- `docs/User_Manual.pdf` — User manual & troubleshooting guide

---

*Dajayana Trading — IT Solutions & Managed Services*
*Lot 7584, Jalan Interhill, Taman Horizon, 98000 Miri, Sarawak, Malaysia*
