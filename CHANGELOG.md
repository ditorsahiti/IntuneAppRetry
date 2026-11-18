# Changelog

## Unreleased
- **Cancel** and **Stop monitoring / cache** merged into one context-aware button (label changes: Cancel scan / Cancel sign-in / Cancel sync / Stop monitoring / cache)
- Removed UTF-8 BOM and added a blank first line so the script can be copied and pasted from GitHub

## 1.3.1
- App names are now also resolved locally from the IME log cache (policy ID/name pairs), without Graph sign-in
- New **Cancel** button: stops a running scan (also the automatic scan on start), Graph sign-in or MDM sync
- New **Cancel sign-in** button in the device code panel
- A running retry cannot be cancelled (protects registry reset and IME restart)

## 1.3.0
- Live monitor (registry + IME log evidence), retry history
- Optional installation cache capture from IMECache
- Graph name resolution via device code flow (optional)
- Registry backups with SHA256 manifest
