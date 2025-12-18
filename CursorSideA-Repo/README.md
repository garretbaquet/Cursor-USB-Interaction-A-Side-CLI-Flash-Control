# Cursor Side-A Utility (PowerShell)

A small PowerShell repo you can run from **Cursor’s integrated terminal** (Windows) to:

- verify/install **Python**, **Git**, **PlatformIO**
- optionally add **Cursor** to your User PATH
- list **USB/COM ports** with best-effort VID/PID hints
- **upload firmware** via `pio run -t upload`
- **monitor serial output**
- **apply Side-A configuration** by sending your Side-A CLI commands over serial

> This repo is intended to *control* your existing Side-A firmware repo (the one with `platformio.ini`), not replace it.

## Files

- `CursorSideA.ps1` — main script
- `sidea-config.example.json` — example config payload
- `.gitignore` — standard ignores

## Quick start

### 1) List ports
```powershell
powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -ListPorts
```

### 2) Upload firmware
```powershell
powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -InstallTools
powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -Upload -RepoPath "C:\path\to\ESP32S3-A-Side"
```

### 3) Apply config + monitor
```powershell
powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -ApplyConfig -Port COM7 -ConfigPath ".\sidea-config.json"
powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -Monitor -Port COM7
```

### One-shot: upload + config + monitor
```powershell
powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -Upload -ApplyConfig -Monitor -RepoPath "C:\path\to\ESP32S3-A-Side" -ConfigPath ".\sidea-config.json"
```

## Interactive menu
```powershell
powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -Menu
```

## Notes

- Serial access generally works without admin, but **driver install** might need admin.
- The config apply step assumes your Side-A firmware supports commands like:
  `node`, `i2c`, `scan`, `probe`, `fmt`, `live`, `rate`, `thr set`, `led`, `sound`, `cfg save`, `cfg show`, `telem`.

## License

Pick whatever you like (MIT/Apache-2.0). This repo contains no Bosch code.
