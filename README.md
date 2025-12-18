# Cursor-USB-Interaction-A-Side-CLI-Flash-Control
Install/verify tooling (Python, PlatformIO, Git) + optionally add Cursor to PATH  Scan USB/COM ports (with VID/PID hints when available)  Flash Side-A firmware via pio run -t upload  Monitor USB serial output (native .NET SerialPort)  Push Side-A CLI configuration (node id, i2c, fmt, live, thresholds, led, sound, rates, cal load/save, etc.)

1) Where to save the script (recommended layout)

Create a folder like:

C:\scripts\CursorSideA\ (simple, no spaces)

Put these files there:

CursorSideA.ps1

sidea-config.json (copy from sidea-config.example.json and edit)

So you end up with:

C:\scripts\CursorSideA\CursorSideA.ps1

C:\scripts\CursorSideA\sidea-config.json

2) Install prerequisites (one-time)
A) Install Git (if needed)

Install Git for Windows (or via winget if you already have it)

If you don’t know: just install Git for Windows normally.

B) Install Python 3.12+ (if needed)

Install Python 3.12 from python.org or Microsoft Store.

Confirm:

py -V

C) Install PlatformIO Core (if needed)
py -m pip install --upgrade pip
py -m pip install --upgrade platformio


Confirm:

pio --version

3) Open the right PowerShell (admin vs non-admin)
When you need Admin

You might need Admin for:

Driver installs (first time a USB-serial chip shows up)

Some winget installs

System PATH changes (machine-wide)

Safest default

Use Admin PowerShell for the first-time setup, then normal PowerShell later.

Open Admin PowerShell:

Start Menu → type PowerShell

Right-click → Run as administrator

4) Allow running the script (Execution Policy)

In Admin PowerShell, run:

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser


That’s the common “scripts work” setting without going full chaos mode.

5) Run the script
A) Go to your script folder
cd C:\scripts\CursorSideA

B) First run: install/verify tools
powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -InstallTools

C) Plug in the ESP32-S3 Side-A board and list ports
powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -ListPorts


You’ll see COM ports listed. Note the one that looks like ESP32 / USB Serial (example: COM7).

6) Flash firmware (PlatformIO upload)

This script uploads firmware from your Side-A firmware repo (the one that has platformio.ini).

Example:

powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 `
  -Upload `
  -RepoPath "C:\path\to\ESP32S3-A-Side" `
  -Port COM7


If you omit -Port, it will try to auto-pick the “best” COM port.

7) Write configuration to Side-A over serial
A) Make your config file

Copy example → real config:

sidea-config.example.json → sidea-config.json

Edit node_id, thresholds, etc.

B) Apply config
powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 `
  -ApplyConfig `
  -Port COM7 `
  -ConfigPath ".\sidea-config.json"

8) Monitor serial output
powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -Monitor -Port COM7


Stop monitor with Ctrl + C.

9) One-shot: Upload + Apply config + Monitor

This is the “do the whole ritual” command:

powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 `
  -Upload -ApplyConfig -Monitor `
  -RepoPath "C:\path\to\ESP32S3-A-Side" `
  -ConfigPath ".\sidea-config.json" `
  -Port COM7

10) Use the interactive menu (best for new users)
powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -Menu


Then pick:

install tools

list ports

upload

apply config

monitor

Common “gotchas” (and fixes)
“Script cannot be loaded because running scripts is disabled…”

Run:

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

“pio not recognized”

Close terminal → open a new PowerShell window (PATH refresh), then:

pio --version

“Access denied” on COM port

Close any other serial monitor (Arduino IDE, PlatformIO monitor, PuTTY, etc.)

Unplug/replug board

Try again

Git asks for merge commit message during pull

If it opens the editor, save+exit:

Vim: Esc then :wq then Enter

Nano: Ctrl+O, Enter, Ctrl+X
