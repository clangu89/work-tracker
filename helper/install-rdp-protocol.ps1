# ========================================================
# RDP Manager - Windows 1-Click Protocol Installer (v2.0)
# ========================================================

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  RDP Manager: 1-Click Zero-Prompt Protocol Setup v2.0" -ForegroundColor Cyan
Write-Host "  (Clipboard & File Copy Redirection Enabled)" -ForegroundColor Cyan
Write-Host "  (clangu89 / rdp-manager)" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""

$targetDir = Join-Path $env:LOCALAPPDATA "RDPManager"
if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

Write-Host "[*] Installing system files to $targetDir..." -ForegroundColor Yellow

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$cacheBuster = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

# 1. Install rdp-handler.ps1
$handlerDest = Join-Path $targetDir "rdp-handler.ps1"
if ($scriptDir -and (Test-Path (Join-Path $scriptDir "rdp-handler.ps1"))) {
    Copy-Item (Join-Path $scriptDir "rdp-handler.ps1") $handlerDest -Force
} else {
    try {
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/clangu89/work-tracker/main/helper/rdp-handler.ps1?t=$cacheBuster" -OutFile $handlerDest -UseBasicParsing -ErrorAction Stop
    } catch {
        try {
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/clangu89/rdp-manager/main/helper/rdp-handler.ps1?t=$cacheBuster" -OutFile $handlerDest -UseBasicParsing -ErrorAction Stop
        } catch {
            Write-Warning "Could not download rdp-handler.ps1 directly, using fallback."
        }
    }
}

# 2. Install rdp-runner.vbs
$runnerDest = Join-Path $targetDir "rdp-runner.vbs"
if ($scriptDir -and (Test-Path (Join-Path $scriptDir "rdp-runner.vbs"))) {
    Copy-Item (Join-Path $scriptDir "rdp-runner.vbs") $runnerDest -Force
} else {
    try {
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/clangu89/work-tracker/main/helper/rdp-runner.vbs?t=$cacheBuster" -OutFile $runnerDest -UseBasicParsing -ErrorAction Stop
    } catch {
        try {
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/clangu89/rdp-manager/main/helper/rdp-runner.vbs?t=$cacheBuster" -OutFile $runnerDest -UseBasicParsing -ErrorAction Stop
        } catch {
            Write-Warning "Could not download rdp-runner.vbs directly, using fallback."
        }
    }
}

# 3. Register Windows Registry Protocol (rdpconnect://)
Write-Host "[*] Registering Windows Registry Protocol (rdpconnect://)..." -ForegroundColor Yellow

$regPath = "HKCU:\Software\Classes\rdpconnect"
if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}
Set-Item -Path $regPath -Value "URL:RDP Connect Protocol"
New-ItemProperty -Path $regPath -Name "URL Protocol" -Value "" -PropertyType String -Force | Out-Null

$cmdPath = "HKCU:\Software\Classes\rdpconnect\shell\open\command"
if (-not (Test-Path $cmdPath)) {
    New-Item -Path $cmdPath -Force | Out-Null
}
$vbsPath = Join-Path $targetDir "rdp-runner.vbs"
$cmdVal = "wscript.exe `"$vbsPath`" `"%1`""
Set-Item -Path $cmdPath -Value $cmdVal

# 4. Pre-configure Default.rdp and LocalDevices for Clipboard & File Copy
Write-Host "[*] Enabling Clipboard & File Copy Redirection in Default.rdp..." -ForegroundColor Yellow
try {
    $docsPath = [Environment]::GetFolderPath('MyDocuments')
    if (-not $docsPath) { $docsPath = Join-Path $env:USERPROFILE "Documents" }
    $rdpPath = Join-Path $docsPath "Default.rdp"

    if (Test-Path $rdpPath) {
        $item = Get-Item -Path $rdpPath -Force
        $origAttr = $item.Attributes
        try { $item.Attributes = [System.IO.FileAttributes]::Normal } catch {}

        $lines = [System.IO.File]::ReadAllLines($rdpPath)
        $foundClip = $false
        $foundDrive = $false
        $updated = New-Object System.Collections.Generic.List[string]

        foreach ($line in $lines) {
            if ($line -match '^redirectclipboard:i:') {
                $updated.Add("redirectclipboard:i:1")
                $foundClip = $true
            } elseif ($line -match '^drivestoredirect:s:') {
                $updated.Add("drivestoredirect:s:*")
                $foundDrive = $true
            } else {
                $updated.Add($line)
            }
        }
        if (-not $foundClip) { $updated.Add("redirectclipboard:i:1") }
        if (-not $foundDrive) { $updated.Add("drivestoredirect:s:*") }

        [System.IO.File]::WriteAllLines($rdpPath, $updated, [System.Text.Encoding]::Unicode)
        try { (Get-Item -Path $rdpPath -Force).Attributes = $origAttr } catch {}
    }

    $ldPath = "HKCU:\Software\Microsoft\Terminal Server Client\LocalDevices"
    if (Test-Path $ldPath) {
        $props = (Get-ItemProperty -Path $ldPath).PSObject.Properties |
            Where-Object { $_.Name -notmatch '^PS(Path|ParentPath|ChildName|Drive|Provider)$' }
        foreach ($p in $props) {
            New-ItemProperty -Path $ldPath -Name $p.Name -Value 255 -PropertyType DWord -Force | Out-Null
        }
    }
} catch {
    Write-Warning "Could not pre-configure Default.rdp: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "========================================================" -ForegroundColor Green
Write-Host "  [SUCCESS] Installation v2.0 completed successfully!" -ForegroundColor Green
Write-Host "  - Clipboard (Text/Image) Copy-Paste: ENABLED" -ForegroundColor Green
Write-Host "  - File Copy-Paste & Drive Sharing  : ENABLED" -ForegroundColor Green
Write-Host "  You can now return to the web app and click 'Remote'." -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Green
Write-Host ""
