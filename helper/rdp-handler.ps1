param([string]$uri)

$HANDLER_VERSION = 2

function Update-DefaultRdpProfile {
    param(
        [bool]$EnableClipboard = $true,
        [bool]$EnableDrives = $true
    )
    try {
        $docsPath = [Environment]::GetFolderPath('MyDocuments')
        if (-not $docsPath) {
            $docsPath = Join-Path $env:USERPROFILE "Documents"
        }
        $rdpPath = Join-Path $docsPath "Default.rdp"
        $clipVal = if ($EnableClipboard) { "1" } else { "0" }
        $driveVal = if ($EnableDrives) { "*" } else { "" }

        if (Test-Path $rdpPath) {
            $item = Get-Item -Path $rdpPath -Force
            $origAttr = $item.Attributes
            try {
                $item.Attributes = [System.IO.FileAttributes]::Normal
            } catch {}

            $lines = [System.IO.File]::ReadAllLines($rdpPath)
            $foundClip = $false
            $foundDrive = $false
            $foundPrompt = $false
            $updated = New-Object System.Collections.Generic.List[string]

            foreach ($line in $lines) {
                if ($line -match '^redirectclipboard:i:') {
                    $updated.Add("redirectclipboard:i:$clipVal")
                    $foundClip = $true
                } elseif ($line -match '^drivestoredirect:s:') {
                    $updated.Add("drivestoredirect:s:$driveVal")
                    $foundDrive = $true
                } elseif ($line -match '^prompt for credentials:i:') {
                    $updated.Add("prompt for credentials:i:0")
                    $foundPrompt = $true
                } else {
                    $updated.Add($line)
                }
            }

            if (-not $foundClip) { $updated.Add("redirectclipboard:i:$clipVal") }
            if (-not $foundDrive) { $updated.Add("drivestoredirect:s:$driveVal") }
            if (-not $foundPrompt) { $updated.Add("prompt for credentials:i:0") }

            [System.IO.File]::WriteAllLines($rdpPath, $updated, [System.Text.Encoding]::Unicode)
            try {
                (Get-Item -Path $rdpPath -Force).Attributes = $origAttr
            } catch {}
        } else {
            $defaultLines = @(
                "screen mode id:i:2",
                "use multimon:i:0",
                "session bpp:i:32",
                "compression:i:1",
                "keyboardhook:i:2",
                "audiomode:i:0",
                "redirectprinters:i:0",
                "redirectcomports:i:0",
                "redirectsmartcards:i:0",
                "redirectclipboard:i:$clipVal",
                "drivestoredirect:s:$driveVal",
                "displayconnectionbar:i:1",
                "autoreconnection enabled:i:1",
                "authentication level:i:2",
                "prompt for credentials:i:0",
                "negotiate security layer:i:1"
            )
            [System.IO.File]::WriteAllLines($rdpPath, $defaultLines, [System.Text.Encoding]::Unicode)
        }
    } catch {
        # Ignore Default.rdp write errors so connection can still proceed
    }
}

function Set-RdpLocalDeviceTrust {
    param(
        [string]$Ip,
        [int]$Port = 3389
    )
    try {
        $ldPath = "HKCU:\Software\Microsoft\Terminal Server Client\LocalDevices"
        if (-not (Test-Path $ldPath)) {
            New-Item -Path $ldPath -Force | Out-Null
        }
        # 255 (0xFF) authorizes all local device redirections (Clipboard, Drives, etc.) without prompt
        New-ItemProperty -Path $ldPath -Name $Ip -Value 255 -PropertyType DWord -Force | Out-Null
        if ($Port -and $Port -ne 3389) {
            New-ItemProperty -Path $ldPath -Name "$Ip`:$Port" -Value 255 -PropertyType DWord -Force | Out-Null
        }
    } catch {}
}

function Invoke-RemoteClipboardFixAutoType {
    param([string]$CmdType = "restart_rdpclip")

    $mstscProc = Get-Process -Name "mstsc" -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne 0 } |
        Select-Object -First 1

    if (-not $mstscProc) {
        $msgB64 = "4LmE4Lih4LmI4Lie4Lia4Lir4LiZ4LmJ4Liy4LiV4LmI4Liy4LiH4LmC4Lib4Lij4LmB4LiB4Lij4LihIFJlbW90ZSBEZXNrdG9wIChtc3RzYy5leGUpIOC4l+C4teC5iOC4geC4s+C4peC4seC4h+C5gOC4m+C4tOC4lOC4reC4ouC4ueC5iA0KDQrguKfguLTguJjguLXguYPguIrguYnguIfguLLguJk6DQoxLiDguYDguJvguLTguJTguKvguJnguYnguLLguJXguYjguLLguIcgUmVtb3RlIERlc2t0b3Ag4LiX4Li14LmI4LiV4LmJ4Lit4LiH4LiB4Liy4Lij4LiL4LmI4Lit4LihDQoyLiDguITguKXguLTguIHguILguKfguLLguJvguLjguYjguKEgU3RhcnQg4LmA4Lil4Li34Lit4LiBIFJ1biAo4Lir4Lij4Li34Lit4LmA4Lib4Li04LiU4Lir4LiZ4LmJ4Liy4LiV4LmI4Liy4LiHIENNRCkg4LiX4Li04LmJ4LiH4LmE4Lin4LmJ4Lia4LiZ4LmA4LiE4Lij4Li34LmI4Lit4LiH4Lib4Lil4Liy4Lii4LiX4Liy4LiHDQozLiDguIHguKXguLHguJrguKHguLLguIHguJTguJvguLjguYjguKHguJnguLXguYnguK3guLXguIHguITguKPguLHguYnguIcg4Lij4Liw4Lia4Lia4LiI4Liw4Liq4Lil4Lix4Lia4LmE4Lib4Lie4Li04Lih4Lie4LmM4LiE4Liz4Liq4Lix4LmI4LiH4LiL4LmI4Lit4Lih4LmD4Lir4LmJ4Lit4Lix4LiV4LmC4LiZ4Lih4Lix4LiV4Li0"
        $thaiMsg = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($msgB64))
        [System.Windows.Forms.MessageBox]::Show(
            $thaiMsg,
            "RDP Manager - Clipboard Fixer",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $csharp = @"
using System;
using System.Runtime.InteropServices;
using System.Threading;

public static class RdpAutoTyper {
    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [StructLayout(LayoutKind.Sequential)]
    public struct INPUT {
        public uint type;
        public InputUnion U;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct InputUnion {
        [FieldOffset(0)] public MOUSEINPUT mi;
        [FieldOffset(0)] public KEYBDINPUT ki;
        [FieldOffset(0)] public HARDWAREINPUT hi;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MOUSEINPUT {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct KEYBDINPUT {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct HARDWAREINPUT {
        public uint uMsg;
        public ushort wParamL;
        public ushort wParamH;
    }

    public const uint INPUT_KEYBOARD = 1;
    public const uint KEYEVENTF_KEYUP = 0x0002;
    public const uint KEYEVENTF_UNICODE = 0x0004;
    public const ushort VK_RETURN = 0x0D;

    public static void FocusAndType(IntPtr hWnd, string text, bool pressEnter) {
        if (IsIconic(hWnd)) {
            ShowWindow(hWnd, 9); // SW_RESTORE
        }
        SetForegroundWindow(hWnd);
        Thread.Sleep(650);

        int size = Marshal.SizeOf(typeof(INPUT));
        foreach (char c in text) {
            INPUT[] inputs = new INPUT[2];
            inputs[0].type = INPUT_KEYBOARD;
            inputs[0].U.ki.wVk = 0;
            inputs[0].U.ki.wScan = (ushort)c;
            inputs[0].U.ki.dwFlags = KEYEVENTF_UNICODE;

            inputs[1].type = INPUT_KEYBOARD;
            inputs[1].U.ki.wVk = 0;
            inputs[1].U.ki.wScan = (ushort)c;
            inputs[1].U.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;

            SendInput(2, inputs, size);
            Thread.Sleep(12);
        }

        if (pressEnter) {
            Thread.Sleep(80);
            INPUT[] enterInputs = new INPUT[2];
            enterInputs[0].type = INPUT_KEYBOARD;
            enterInputs[0].U.ki.wVk = VK_RETURN;
            enterInputs[0].U.ki.dwFlags = 0;

            enterInputs[1].type = INPUT_KEYBOARD;
            enterInputs[1].U.ki.wVk = VK_RETURN;
            enterInputs[1].U.ki.dwFlags = KEYEVENTF_KEYUP;

            SendInput(2, enterInputs, size);
        }
    }
}
"@

    Add-Type -TypeDefinition $csharp -ErrorAction SilentlyContinue

    $cmdToType = 'cmd /c "taskkill /f /im rdpclip.exe & start rdpclip.exe"'
    if ($CmdType -eq "unlock_policy") {
        $cmdToType = 'reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" /v fDisableClip /t REG_DWORD /d 0 /f & reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" /v fDisableCdm /t REG_DWORD /d 0 /f & taskkill /f /im rdpclip.exe & start rdpclip.exe'
    }

    [RdpAutoTyper]::FocusAndType($mstscProc.MainWindowHandle, $cmdToType, $true)
}

function Start-BackgroundSelfUpdate {
    param([int]$RemoteVersion)
    try {
        if ($RemoteVersion -le $HANDLER_VERSION) { return }
        $selfPath = $MyInvocation.PSCommandPath
        if (-not $selfPath) {
            $selfPath = Join-Path (Join-Path $env:LOCALAPPDATA "RDPManager") "rdp-handler.ps1"
        }
        $updateCmd = @"
Start-Sleep -Seconds 2
try {
    `$url = 'https://raw.githubusercontent.com/clangu89/rdp-manager/main/helper/rdp-handler.ps1?v=$RemoteVersion'
    `$tmp = '$selfPath.tmp'
    Invoke-WebRequest -Uri `$url -OutFile `$tmp -UseBasicParsing -ErrorAction Stop
    if ((Test-Path `$tmp) -and ((Get-Content `$tmp -Raw) -match 'HANDLER_VERSION')) {
        Move-Item -Path `$tmp -Destination '$selfPath' -Force
    }
} catch {}
"@
        Start-Process "powershell.exe" -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"$updateCmd`""
    } catch {}
}

try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

    if (-not $uri) {
        exit
    }

    # Remove rdpconnect:// prefix and trailing slash
    $raw = $uri -replace '^rdpconnect://', '' -replace '/$', ''
    $raw = [System.Uri]::UnescapeDataString($raw)

    # Decode Base64 JSON
    $jsonBytes = [System.Convert]::FromBase64String($raw)
    $jsonText = [System.Text.Encoding]::UTF8.GetString($jsonBytes)
    $data = $jsonText | ConvertFrom-Json

    # Check if Auto-Type Fix Remote Clipboard mode is requested
    if ($data.action -eq "fix_clipboard") {
        $cmdType = if ($data.cmdType) { [string]$data.cmdType } else { "restart_rdpclip" }
        Invoke-RemoteClipboardFixAutoType -CmdType $cmdType
        if ($data.v) { Start-BackgroundSelfUpdate -RemoteVersion ([int]$data.v) }
        exit
    }

    $ip = $data.ip
    $port = if ($data.port) { [int]$data.port } else { 3389 }
    $user = $data.user
    $pass = $data.pass
    $enableClipboard = if ($null -ne $data.clipboard) { [bool]$data.clipboard } else { $true }
    $enableDrives = if ($null -ne $data.drives) { [bool]$data.drives } else { $true }

    if ($ip -and $user) {
        # 1. Configure Default.rdp for Clipboard & File/Drive Redirection
        Update-DefaultRdpProfile -EnableClipboard $enableClipboard -EnableDrives $enableDrives

        # 2. Pre-authorize LocalDevices in Registry (0xFF) so no warning dialog blocks Drives/Clipboard
        Set-RdpLocalDeviceTrust -Ip $ip -Port $port

        # 3. Store credential in Windows Credential Manager (both base IP and IP:Port for custom ports)
        cmdkey /generic:"TERMSRV/$ip" /user:$user /pass:$pass | Out-Null
        if ($port -ne 3389) {
            cmdkey /generic:"TERMSRV/$ip`:$port" /user:$user /pass:$pass | Out-Null
        }

        # 4. Launch Windows Remote Desktop Connection
        Start-Process "mstsc.exe" -ArgumentList "/v:$ip`:$port"
    }

    if ($data.v) {
        Start-BackgroundSelfUpdate -RemoteVersion ([int]$data.v)
    }
} catch {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    [System.Windows.Forms.MessageBox]::Show("Failed to connect Remote Desktop: " + $_.Exception.Message, "RDP Manager Error", 0, 16)
}
