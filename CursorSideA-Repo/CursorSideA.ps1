#Requires -Version 5.1
<#
CursorSideA.ps1
- Flash firmware (PlatformIO)
- Scan USB/COM ports
- Monitor USB serial output
- Write Side-A configuration via serial CLI commands

Examples:
  # List ports
  powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -ListPorts

  # Upload firmware (auto-detect port, assumes repo contains platformio.ini)
  powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -Upload -RepoPath "C:\path\to\ESP32S3-A-Side"

  # Monitor serial
  powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -Monitor -Port COM7

  # Apply config JSON to device
  powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -ApplyConfig -Port COM7 -ConfigPath ".\sidea-config.json"

  # Upload + apply config + start monitor
  powershell -ExecutionPolicy Bypass -File .\CursorSideA.ps1 -Upload -ApplyConfig -Monitor -RepoPath "C:\path\to\repo" -ConfigPath ".\sidea-config.json"

Notes:
- “Give Cursor system access” here means: ensure tools exist + PATH is set so Cursor’s integrated terminal can call them.
- Serial/COM access on Windows usually does NOT require admin; driver installation might.
#>

[CmdletBinding()]
param(
  [switch]$InstallTools,
  [switch]$AddCursorToPath,
  [switch]$ListPorts,
  [switch]$Upload,
  [switch]$Monitor,
  [switch]$ApplyConfig,
  [switch]$Menu,

  [string]$RepoPath = "",
  [string]$Env = "",

  [string]$Port = "",
  [int]$Baud = 115200,

  [string]$ConfigPath = "",

  [int]$UploadTimeoutSec = 600,
  [int]$SerialReadTimeoutMs = 250,
  [int]$CmdTimeoutMs = 1500,

  [switch]$CalLoad,
  [switch]$CalSave
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# -----------------------------
# Logging helpers
# -----------------------------
function Write-Log {
  param(
    [Parameter(Mandatory=$true)][string]$Message,
    [ValidateSet("INFO","WARN","ERROR","OK","DBG")][string]$Level="INFO"
  )
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $prefix = "[$ts][$Level]"
  switch ($Level) {
    "ERROR" { Write-Host "$prefix $Message" -ForegroundColor Red }
    "WARN"  { Write-Host "$prefix $Message" -ForegroundColor Yellow }
    "OK"    { Write-Host "$prefix $Message" -ForegroundColor Green }
    "DBG"   { Write-Host "$prefix $Message" -ForegroundColor DarkGray }
    default { Write-Host "$prefix $Message" }
  }
}

function Test-Admin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-Command {
  param([Parameter(Mandatory=$true)][string]$Name)
  $cmd = Get-Command $Name -ErrorAction SilentlyContinue
  return [bool]$cmd
}

function Get-OsArch {
  if ([Environment]::Is64BitOperatingSystem) { return "x64" }
  return "x86"
}

# -----------------------------
# Tooling install / verification
# -----------------------------
function Ensure-Winget {
  if (Test-Command "winget") {
    Write-Log "winget found." "OK"
    return
  }
  Write-Log "winget not found. Install 'App Installer' from Microsoft Store (or ensure winget is available) then re-run." "ERROR"
  throw "winget missing"
}

function Ensure-Python {
  if (Test-Command "py") {
    $ver = & py -V 2>&1
    Write-Log "Python launcher found: $ver" "OK"
    return
  }
  Ensure-Winget
  Write-Log "Installing Python (winget)..." "INFO"
  $candidates = @(
    "Python.Python.3.12",
    "Python.Python.3.11",
    "Python.Python.3"
  )
  $installed = $false
  foreach ($id in $candidates) {
    try {
      & winget install --id $id -e --silent --accept-package-agreements --accept-source-agreements | Out-Null
      $installed = $true
      break
    } catch {
      Write-Log "winget install failed for $id (trying next)..." "WARN"
    }
  }
  if (-not $installed) { throw "Python install failed via winget. Install Python 3.12+ manually and re-run." }
  if (-not (Test-Command "py")) { throw "Python installed but 'py' not found. Open a NEW terminal (PATH refresh) and retry." }
  Write-Log "Python installed and launcher is available." "OK"
}

function Ensure-Git {
  if (Test-Command "git") {
    $ver = & git --version 2>&1
    Write-Log "Git found: $ver" "OK"
    return
  }
  Ensure-Winget
  Write-Log "Installing Git (winget)..." "INFO"
  & winget install --id Git.Git -e --silent --accept-package-agreements --accept-source-agreements | Out-Null
  if (-not (Test-Command "git")) { throw "Git install failed or PATH not updated. Open a NEW terminal and retry." }
  Write-Log "Git installed." "OK"
}

function Ensure-PlatformIO {
  if (Test-Command "pio") {
    $ver = & pio --version 2>&1
    Write-Log "PlatformIO found: $ver" "OK"
    return
  }
  Ensure-Python
  Write-Log "Installing PlatformIO Core via pip..." "INFO"
  # Prefer py -m pip to avoid PATH weirdness
  & py -m pip install --upgrade pip | Out-Null
  & py -m pip install --upgrade platformio | Out-Null

  # Refresh current process PATH (pip installs user scripts in AppData\Roaming\Python\Python3x\Scripts)
  $userScripts = Join-Path $env:APPDATA "Python\Python312\Scripts"
  if (Test-Path $userScripts) {
    if (-not ($env:Path -split ';' | Where-Object { $_ -ieq $userScripts })) {
      $env:Path = "$userScripts;$env:Path"
      Write-Log "Temporarily added to PATH for this session: $userScripts" "DBG"
    }
  }

  if (-not (Test-Command "pio")) { throw "PlatformIO install succeeded but 'pio' not found. Open a NEW terminal and retry." }
  Write-Log "PlatformIO installed." "OK"
}

function Try-AddCursorToPath {
  # Typical Cursor install location (may vary)
  $candidates = @(
    (Join-Path $env:LOCALAPPDATA "Programs\Cursor\Cursor.exe"),
    (Join-Path $env:ProgramFiles "Cursor\Cursor.exe"),
    (Join-Path ${env:ProgramFiles(x86)} "Cursor\Cursor.exe")
  ) | Where-Object { $_ -and (Test-Path $_) }

  if (-not $candidates -or $candidates.Count -eq 0) {
    Write-Log "Cursor.exe not found in common locations. Skipping PATH update." "WARN"
    return
  }

  $cursorExe = $candidates[0]
  $cursorDir = Split-Path -Parent $cursorExe
  Write-Log "Found Cursor: $cursorExe" "OK"

  $currentUserPath = [Environment]::GetEnvironmentVariable("Path", "User")
  if (-not $currentUserPath) { $currentUserPath = "" }

  $parts = $currentUserPath -split ';' | Where-Object { $_ -and $_.Trim().Length -gt 0 }
  $already = $parts | Where-Object { $_ -ieq $cursorDir }
  if ($already) {
    Write-Log "Cursor directory already in User PATH." "OK"
    return
  }

  $newUserPath = ($parts + $cursorDir) -join ';'
  [Environment]::SetEnvironmentVariable("Path", $newUserPath, "User")
  Write-Log "Added Cursor directory to User PATH. Open a NEW terminal to pick it up." "OK"
}

# -----------------------------
# Port enumeration (USB/COM)
# -----------------------------
function Get-SerialPorts {
  $ports = @()

  # Try CIM/WMI Win32_SerialPort (friendly names)
  try {
    $wmi = Get-CimInstance Win32_SerialPort -ErrorAction Stop
    foreach ($p in $wmi) {
      $com = $p.DeviceID
      $name = $p.Name
      $pnp = $p.PNPDeviceID
      $vid = $null; $pid = $null
      if ($pnp -match "VID_([0-9A-Fa-f]{4})") { $vid = $Matches[1].ToUpper() }
      if ($pnp -match "PID_([0-9A-Fa-f]{4})") { $pid = $Matches[1].ToUpper() }
      $ports += [pscustomobject]@{
        Port = $com
        Name = $name
        PnpId = $pnp
        Vid = $vid
        Pid = $pid
        Source = "Win32_SerialPort"
      }
    }
  } catch {
    Write-Log "Win32_SerialPort query failed; falling back to [System.IO.Ports.SerialPort]::GetPortNames()" "WARN"
  }

  if (-not $ports -or $ports.Count -eq 0) {
    $names = [System.IO.Ports.SerialPort]::GetPortNames() | Sort-Object
    foreach ($n in $names) {
      $ports += [pscustomobject]@{
        Port = $n
        Name = $n
        PnpId = $null
        Vid = $null
        Pid = $null
        Source = "SerialPortNames"
      }
    }
  }

  return $ports | Sort-Object Port
}

function Choose-BestPort {
  param([Parameter(Mandatory=$true)]$Ports)

  # Heuristics: prefer likely ESP32 / USB-serial devices
  $keywords = @("ESP32","USB JTAG","CP210","Silicon Labs","CH340","CH910","USB Serial","Espressif")
  foreach ($k in $keywords) {
    $hit = $Ports | Where-Object { $_.Name -and $_.Name.ToString().IndexOf($k, [StringComparison]::OrdinalIgnoreCase) -ge 0 } | Select-Object -First 1
    if ($hit) { return $hit.Port }
  }

  # Prefer Espressif VID if present (303A is common for native USB; not universal)
  $vidHit = $Ports | Where-Object { $_.Vid -eq "303A" } | Select-Object -First 1
  if ($vidHit) { return $vidHit.Port }

  # Otherwise first port
  return ($Ports | Select-Object -First 1).Port
}

function Print-Ports {
  $ports = Get-SerialPorts
  if (-not $ports -or $ports.Count -eq 0) {
    Write-Log "No COM ports found." "WARN"
    return
  }

  Write-Host ""
  Write-Host "Detected Serial Ports:" -ForegroundColor Cyan
  Write-Host "---------------------" -ForegroundColor Cyan
  foreach ($p in $ports) {
    $vp = ""
    if ($p.Vid -or $p.Pid) { $vp = " VID=$($p.Vid) PID=$($p.Pid)" }
    Write-Host ("{0,-6} {1}{2}" -f $p.Port, $p.Name, $vp)
  }

  $best = Choose-BestPort -Ports $ports
  Write-Host ""
  Write-Log "Auto-selected port (heuristic): $best" "INFO"
}

# -----------------------------
# PlatformIO upload
# -----------------------------
function Resolve-RepoPath {
  param([string]$PathHint)

  if ($PathHint -and (Test-Path $PathHint)) { return (Resolve-Path $PathHint).Path }

  # Try current directory
  $cwd = (Get-Location).Path
  if (Test-Path (Join-Path $cwd "platformio.ini")) { return $cwd }

  # Try common repo folder names under user profile
  $guessRoots = @(
    (Join-Path $env:USERPROFILE "Documents"),
    (Join-Path $env:USERPROFILE "source"),
    (Join-Path $env:USERPROFILE "repos"),
    (Join-Path $env:USERPROFILE "OneDrive\Documents")
  ) | Where-Object { $_ -and (Test-Path $_) }

  foreach ($r in $guessRoots) {
    try {
      $hit = Get-ChildItem -Path $r -Recurse -Filter "platformio.ini" -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match "ESP32S3-A-Side|ESP32S3.*Side|Side-A|SideA|ESP32AB" } |
        Select-Object -First 1
      if ($hit) { return (Split-Path -Parent $hit.FullName) }
    } catch { }
  }

  throw "RepoPath not found. Provide -RepoPath pointing to folder containing platformio.ini."
}

function Upload-Firmware {
  param(
    [Parameter(Mandatory=$true)][string]$Repo,
    [string]$UploadPort,
    [string]$EnvName,
    [int]$TimeoutSec
  )

  Ensure-PlatformIO

  if (-not (Test-Path (Join-Path $Repo "platformio.ini"))) {
    throw "platformio.ini not found in repo path: $Repo"
  }

  $cmd = @("pio","run","-t","upload")
  if ($EnvName -and $EnvName.Trim().Length -gt 0) {
    $cmd += @("-e",$EnvName)
  }
  if ($UploadPort -and $UploadPort.Trim().Length -gt 0) {
    $cmd += @("--upload-port",$UploadPort)
  }

  Write-Log "Uploading firmware via PlatformIO..." "INFO"
  Write-Log ("Repo: {0}" -f $Repo) "DBG"
  Write-Log ("Cmd : {0}" -f ($cmd -join ' ')) "DBG"

  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $cmd[0]
  $psi.Arguments = ($cmd[1..($cmd.Count-1)] -join ' ')
  $psi.WorkingDirectory = $Repo
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError = $true
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true

  $p = New-Object System.Diagnostics.Process
  $p.StartInfo = $psi
  [void]$p.Start()

  $sw = [Diagnostics.Stopwatch]::StartNew()
  while (-not $p.HasExited) {
    Start-Sleep -Milliseconds 200
    if ($sw.Elapsed.TotalSeconds -gt $TimeoutSec) {
      try { $p.Kill() } catch { }
      throw "Upload timed out after $TimeoutSec sec."
    }
    while (-not $p.StandardOutput.EndOfStream) {
      $line = $p.StandardOutput.ReadLine()
      if ($line -ne $null) { Write-Host $line }
    }
    while (-not $p.StandardError.EndOfStream) {
      $eline = $p.StandardError.ReadLine()
      if ($eline -ne $null) { Write-Host $eline -ForegroundColor DarkRed }
    }
  }

  # Drain remaining output
  $out = $p.StandardOutput.ReadToEnd()
  $err = $p.StandardError.ReadToEnd()
  if ($out) { Write-Host $out }
  if ($err) { Write-Host $err -ForegroundColor DarkRed }

  if ($p.ExitCode -ne 0) {
    throw "PlatformIO upload failed (exit code $($p.ExitCode))."
  }

  Write-Log "Upload complete." "OK"
}

# -----------------------------
# Serial monitor + command sender
# -----------------------------
function Open-SerialPort {
  param(
    [Parameter(Mandatory=$true)][string]$ComPort,
    [int]$BaudRate = 115200,
    [int]$ReadTimeoutMs = 250
  )

  $sp = New-Object System.IO.Ports.SerialPort
  $sp.PortName = $ComPort
  $sp.BaudRate = $BaudRate
  $sp.Parity = [System.IO.Ports.Parity]::None
  $sp.DataBits = 8
  $sp.StopBits = [System.IO.Ports.StopBits]::One
  $sp.Handshake = [System.IO.Ports.Handshake]::None
  $sp.NewLine = "`n"
  $sp.ReadTimeout = $ReadTimeoutMs
  $sp.WriteTimeout = 1000
  $sp.DtrEnable = $true
  $sp.RtsEnable = $true

  $sp.Open()
  return $sp
}

function Monitor-Serial {
  param(
    [Parameter(Mandatory=$true)][string]$ComPort,
    [int]$BaudRate,
    [int]$ReadTimeoutMs
  )

  Write-Log "Starting serial monitor on $ComPort @ $BaudRate (Ctrl+C to stop)..." "INFO"
  $sp = $null
  try {
    $sp = Open-SerialPort -ComPort $ComPort -BaudRate $BaudRate -ReadTimeoutMs $ReadTimeoutMs
    while ($true) {
      try {
        $data = $sp.ReadExisting()
        if ($data -and $data.Length -gt 0) {
          # Print raw, no extra formatting
          Write-Host -NoNewline $data
        } else {
          Start-Sleep -Milliseconds 30
        }
      } catch {
        Start-Sleep -Milliseconds 50
      }
    }
  } finally {
    if ($sp -and $sp.IsOpen) { $sp.Close() }
  }
}

function Send-SerialLine {
  param(
    [Parameter(Mandatory=$true)][System.IO.Ports.SerialPort]$Sp,
    [Parameter(Mandatory=$true)][string]$Line
  )
  $Sp.Write($Line + "`n")
}

function Read-ForDuration {
  param(
    [Parameter(Mandatory=$true)][System.IO.Ports.SerialPort]$Sp,
    [int]$Ms = 400
  )
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $buf = New-Object System.Text.StringBuilder
  while ($sw.ElapsedMilliseconds -lt $Ms) {
    try {
      $chunk = $Sp.ReadExisting()
      if ($chunk) { [void]$buf.Append($chunk) }
    } catch { }
    Start-Sleep -Milliseconds 20
  }
  return $buf.ToString()
}

function Send-CommandAndRead {
  param(
    [Parameter(Mandatory=$true)][System.IO.Ports.SerialPort]$Sp,
    [Parameter(Mandatory=$true)][string]$Cmd,
    [int]$TimeoutMs = 1500
  )

  Write-Log ">> $Cmd" "DBG"
  Send-SerialLine -Sp $Sp -Line $Cmd

  $sw = [Diagnostics.Stopwatch]::StartNew()
  $buf = New-Object System.Text.StringBuilder
  while ($sw.ElapsedMilliseconds -lt $TimeoutMs) {
    try {
      $chunk = $Sp.ReadExisting()
      if ($chunk) { [void]$buf.Append($chunk) }
    } catch { }
    Start-Sleep -Milliseconds 25
  }

  $text = $buf.ToString()
  if ($text -and $text.Trim().Length -gt 0) {
    # Print device response (best-effort)
    Write-Host $text
  }
  return $text
}

# -----------------------------
# Config application (Side-A CLI)
# -----------------------------
function Load-ConfigFile {
  param([Parameter(Mandatory=$true)][string]$Path)

  if (-not (Test-Path $Path)) { throw "ConfigPath not found: $Path" }
  $raw = Get-Content -Path $Path -Raw
  if (-not $raw -or $raw.Trim().Length -eq 0) { throw "Config file is empty: $Path" }

  try {
    return ($raw | ConvertFrom-Json -ErrorAction Stop)
  } catch {
    throw "Config JSON parse failed: $($_.Exception.Message)"
  }
}

function Apply-SideAConfig {
  param(
    [Parameter(Mandatory=$true)][string]$ComPort,
    [int]$BaudRate,
    [int]$ReadTimeoutMs,
    [int]$CmdTimeoutMs,
    [Parameter(Mandatory=$true)]$Cfg,
    [switch]$DoCalLoad,
    [switch]$DoCalSave
  )

  Write-Log "Applying Side-A config on $ComPort..." "INFO"
  $sp = $null

  try {
    $sp = Open-SerialPort -ComPort $ComPort -BaudRate $BaudRate -ReadTimeoutMs $ReadTimeoutMs

    # Give it a moment to dump boot text
    [void](Read-ForDuration -Sp $sp -Ms 600)

    # Optional banner/status first
    Send-CommandAndRead -Sp $sp -Cmd "poster" -TimeoutMs $CmdTimeoutMs | Out-Null
    Send-CommandAndRead -Sp $sp -Cmd "status" -TimeoutMs $CmdTimeoutMs | Out-Null

    # node_id
    if ($Cfg.node_id) {
      Send-CommandAndRead -Sp $sp -Cmd ("node " + $Cfg.node_id) -TimeoutMs $CmdTimeoutMs | Out-Null
    }

    # I2C pins/speed
    if ($Cfg.i2c) {
      $sda = $Cfg.i2c.sda
      $scl = $Cfg.i2c.scl
      $hz  = $Cfg.i2c.hz
      if ($sda -ne $null -and $scl -ne $null) {
        if ($hz -ne $null) {
          Send-CommandAndRead -Sp $sp -Cmd ("i2c {0} {1} {2}" -f $sda,$scl,$hz) -TimeoutMs $CmdTimeoutMs | Out-Null
        } else {
          Send-CommandAndRead -Sp $sp -Cmd ("i2c {0} {1}" -f $sda,$scl) -TimeoutMs $CmdTimeoutMs | Out-Null
        }
        Send-CommandAndRead -Sp $sp -Cmd "scan" -TimeoutMs $CmdTimeoutMs | Out-Null
        Send-CommandAndRead -Sp $sp -Cmd "probe" -TimeoutMs $CmdTimeoutMs | Out-Null
      }
    }

    # Output format
    if ($Cfg.fmt) {
      Send-CommandAndRead -Sp $sp -Cmd ("fmt " + $Cfg.fmt) -TimeoutMs $CmdTimeoutMs | Out-Null
    }
    if ($Cfg.json_pretty -ne $null) {
      Send-CommandAndRead -Sp $sp -Cmd ("json pretty " + [int]$Cfg.json_pretty) -TimeoutMs $CmdTimeoutMs | Out-Null
    }

    # Live interval
    if ($Cfg.live_ms -ne $null) {
      Send-CommandAndRead -Sp $sp -Cmd ("live " + [int]$Cfg.live_ms) -TimeoutMs $CmdTimeoutMs | Out-Null
    }

    # Rates
    if ($Cfg.rates) {
      if ($Cfg.rates.amb) { Send-CommandAndRead -Sp $sp -Cmd ("rate amb " + $Cfg.rates.amb) -TimeoutMs $CmdTimeoutMs | Out-Null }
      if ($Cfg.rates.env) { Send-CommandAndRead -Sp $sp -Cmd ("rate env " + $Cfg.rates.env) -TimeoutMs $CmdTimeoutMs | Out-Null }
      Send-CommandAndRead -Sp $sp -Cmd "probe" -TimeoutMs $CmdTimeoutMs | Out-Null
    }

    # Thresholds: thr set <metric> <warn> <crit>
    if ($Cfg.thresholds) {
      foreach ($metric in @("iaq","co2","voc","temp","rh")) {
        $m = $Cfg.thresholds.$metric
        if ($m -and $m.warn -ne $null -and $m.crit -ne $null) {
          Send-CommandAndRead -Sp $sp -Cmd ("thr set {0} {1} {2}" -f $metric,$m.warn,$m.crit) -TimeoutMs $CmdTimeoutMs | Out-Null
        }
      }
      Send-CommandAndRead -Sp $sp -Cmd "thr show" -TimeoutMs $CmdTimeoutMs | Out-Null
    }

    # LED
    if ($Cfg.led) {
      if ($Cfg.led.mode) {
        Send-CommandAndRead -Sp $sp -Cmd ("led mode " + $Cfg.led.mode) -TimeoutMs $CmdTimeoutMs | Out-Null
      }
      if ($Cfg.led.bright -ne $null) {
        Send-CommandAndRead -Sp $sp -Cmd ("led bright " + [int]$Cfg.led.bright) -TimeoutMs $CmdTimeoutMs | Out-Null
      }
      if ($Cfg.led.rgb -and $Cfg.led.rgb.Count -ge 3) {
        $r = [int]$Cfg.led.rgb[0]; $g = [int]$Cfg.led.rgb[1]; $b = [int]$Cfg.led.rgb[2]
        Send-CommandAndRead -Sp $sp -Cmd ("led rgb {0} {1} {2}" -f $r,$g,$b) -TimeoutMs $CmdTimeoutMs | Out-Null
      }
    }

    # Sound
    if ($Cfg.sound -ne $null) {
      Send-CommandAndRead -Sp $sp -Cmd ("sound " + [int]$Cfg.sound) -TimeoutMs $CmdTimeoutMs | Out-Null
    }

    # Calibration
    if ($DoCalLoad) { Send-CommandAndRead -Sp $sp -Cmd "cal load" -TimeoutMs $CmdTimeoutMs | Out-Null }
    if ($DoCalSave) { Send-CommandAndRead -Sp $sp -Cmd "cal save" -TimeoutMs $CmdTimeoutMs | Out-Null }

    # Force save + show
    Send-CommandAndRead -Sp $sp -Cmd "cfg save" -TimeoutMs $CmdTimeoutMs | Out-Null
    Send-CommandAndRead -Sp $sp -Cmd "cfg show" -TimeoutMs $CmdTimeoutMs | Out-Null
    Send-CommandAndRead -Sp $sp -Cmd "status" -TimeoutMs $CmdTimeoutMs | Out-Null
    Send-CommandAndRead -Sp $sp -Cmd "telem" -TimeoutMs $CmdTimeoutMs | Out-Null

    Write-Log "Config applied." "OK"
  }
  finally {
    if ($sp -and $sp.IsOpen) { $sp.Close() }
  }
}

# -----------------------------
# Menu
# -----------------------------
function Show-Menu {
  Write-Host ""
  Write-Host "==============================" -ForegroundColor Cyan
  Write-Host " Cursor Side-A Utility" -ForegroundColor Cyan
  Write-Host "==============================" -ForegroundColor Cyan
  Write-Host " 1) Install/Verify Tools (Python, Git, PlatformIO)"
  Write-Host " 2) Add Cursor to PATH (User)"
  Write-Host " 3) List Serial Ports"
  Write-Host " 4) Upload Firmware (PlatformIO)"
  Write-Host " 5) Monitor Serial"
  Write-Host " 6) Apply Config JSON"
  Write-Host " 7) Upload + Apply Config + Monitor"
  Write-Host " 0) Exit"
  Write-Host ""
}

function Prompt-Value {
  param([string]$Label, [string]$Default="")
  if ($Default -and $Default.Trim().Length -gt 0) {
    $v = Read-Host "$Label [$Default]"
    if (-not $v -or $v.Trim().Length -eq 0) { return $Default }
    return $v
  } else {
    return (Read-Host "$Label")
  }
}

# -----------------------------
# Main
# -----------------------------
try {
  if ($Menu) {
    while ($true) {
      Show-Menu
      $choice = Read-Host "Choose"
      switch ($choice) {
        "1" {
          Write-Log ("Admin: " + (Test-Admin)) "INFO"
          Ensure-Python
          Ensure-Git
          Ensure-PlatformIO
        }
        "2" { Try-AddCursorToPath }
        "3" { Print-Ports }
        "4" {
          $rp = Prompt-Value "Repo path (folder containing platformio.ini)" $RepoPath
          $rp = Resolve-RepoPath -PathHint $rp
          $ports = Get-SerialPorts
          $best = Choose-BestPort -Ports $ports
          $p = Prompt-Value "Upload port (COMx)" $best
          $e = Prompt-Value "PIO environment (-e), blank to omit" $Env
          Upload-Firmware -Repo $rp -UploadPort $p -EnvName $e -TimeoutSec $UploadTimeoutSec
        }
        "5" {
          $ports = Get-SerialPorts
          $best = Choose-BestPort -Ports $ports
          $p = Prompt-Value "Monitor port (COMx)" $best
          $b = Prompt-Value "Baud" "$Baud"
          Monitor-Serial -ComPort $p -BaudRate ([int]$b) -ReadTimeoutMs $SerialReadTimeoutMs
        }
        "6" {
          $ports = Get-SerialPorts
          $best = Choose-BestPort -Ports $ports
          $p = Prompt-Value "Device port (COMx)" $best
          $cp = Prompt-Value "Config JSON path" $ConfigPath
          $cfg = Load-ConfigFile -Path $cp
          Apply-SideAConfig -ComPort $p -BaudRate $Baud -ReadTimeoutMs $SerialReadTimeoutMs -CmdTimeoutMs $CmdTimeoutMs -Cfg $cfg -DoCalLoad:$CalLoad -DoCalSave:$CalSave
        }
        "7" {
          $rp = Prompt-Value "Repo path (folder containing platformio.ini)" $RepoPath
          $rp = Resolve-RepoPath -PathHint $rp
          $ports = Get-SerialPorts
          $best = Choose-BestPort -Ports $ports
          $p = Prompt-Value "Port (COMx)" $best
          $e = Prompt-Value "PIO environment (-e), blank to omit" $Env
          $cp = Prompt-Value "Config JSON path" $ConfigPath
          $cfg = Load-ConfigFile -Path $cp
          Upload-Firmware -Repo $rp -UploadPort $p -EnvName $e -TimeoutSec $UploadTimeoutSec
          Start-Sleep -Seconds 2
          Apply-SideAConfig -ComPort $p -BaudRate $Baud -ReadTimeoutMs $SerialReadTimeoutMs -CmdTimeoutMs $CmdTimeoutMs -Cfg $cfg -DoCalLoad:$CalLoad -DoCalSave:$CalSave
          Start-Sleep -Milliseconds 500
          Monitor-Serial -ComPort $p -BaudRate $Baud -ReadTimeoutMs $SerialReadTimeoutMs
        }
        "0" { break }
        default { Write-Log "Invalid choice." "WARN" }
      }
    }
    exit 0
  }

  # Non-menu execution
  if ($InstallTools) {
    Write-Log ("Admin: " + (Test-Admin)) "INFO"
    Ensure-Python
    Ensure-Git
    Ensure-PlatformIO
  }

  if ($AddCursorToPath) {
    Try-AddCursorToPath
  }

  if ($ListPorts) {
    Print-Ports
  }

  # Resolve port if needed
  $resolvedPort = $Port
  if (($Upload -or $Monitor -or $ApplyConfig) -and (-not $resolvedPort -or $resolvedPort.Trim().Length -eq 0)) {
    $ports = Get-SerialPorts
    if (-not $ports -or $ports.Count -eq 0) { throw "No COM ports found. Plug in device and retry." }
    $resolvedPort = Choose-BestPort -Ports $ports
    Write-Log "Auto-selected port: $resolvedPort" "INFO"
  }

  if ($Upload) {
    $rp = Resolve-RepoPath -PathHint $RepoPath
    Upload-Firmware -Repo $rp -UploadPort $resolvedPort -EnvName $Env -TimeoutSec $UploadTimeoutSec
  }

  if ($ApplyConfig) {
    if (-not $ConfigPath -or $ConfigPath.Trim().Length -eq 0) { throw "Provide -ConfigPath with Side-A config JSON." }
    $cfg = Load-ConfigFile -Path $ConfigPath
    Apply-SideAConfig -ComPort $resolvedPort -BaudRate $Baud -ReadTimeoutMs $SerialReadTimeoutMs -CmdTimeoutMs $CmdTimeoutMs -Cfg $cfg -DoCalLoad:$CalLoad -DoCalSave:$CalSave
  }

  if ($Monitor) {
    Monitor-Serial -ComPort $resolvedPort -BaudRate $Baud -ReadTimeoutMs $SerialReadTimeoutMs
  }

  if (-not ($InstallTools -or $AddCursorToPath -or $ListPorts -or $Upload -or $ApplyConfig -or $Monitor -or $Menu)) {
    Write-Log "No action specified. Use -Menu for interactive mode, or -ListPorts/-Upload/-Monitor/-ApplyConfig." "WARN"
  }
}
catch {
  Write-Log $_.Exception.Message "ERROR"
  Write-Log "Tip: run -Menu for interactive mode, and use -ListPorts to confirm COM port." "INFO"
  exit 1
}
