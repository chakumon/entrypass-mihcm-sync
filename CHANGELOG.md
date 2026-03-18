# Changelog

All notable changes to EntryPassSync will be documented in this file.

## [1.1.1] - 2026-03-19

### Added
- Dajayana logo on window title bar and system tray icon
- Sync runs immediately on app startup

## [1.1.0] - 2026-03-19

### Added
- Sync on startup -- app now syncs immediately when launched, then every 15 minutes
- Auto-update system -- checks for updates on startup and via About panel
- "Check for Updates" button in About panel with status feedback
- Button changes to "Update Now" (green) when update is available
- CHANGELOG.md for version tracking

### Changed
- Replaced Task Scheduler with built-in 15-minute sync timer
- Auto-start via HKLM\Run (machine-wide, works for any user)
- "Save & Install Schedule" button simplified to "Save Configuration"
- Dashboard status shows "Database Direct Mode" / "File Mode"
- About panel version now dynamic from $script:appVersion

### Removed
- Task Scheduler integration (replaced by built-in timer)
- Schedule section from config panel
- Firebird Username, Password, FB Client DLL from config panel (pre-configured)
- Skipped counter from dashboard (internal only)
- Redundant "Status" label from dashboard

## [1.0.0] - 2026-03-18

### Added
- Initial release
- WinForms GUI with Dashboard, Config, Logs, About panels
- Database Direct Mode (reads TRANS.FDB via Firebird)
- File Mode (reads DATA*.txt export files)
- MiHCM API integration with batch upload
- License validation via GitHub
- System tray with balloon notifications
- Duplicate detection and retry with backoff
