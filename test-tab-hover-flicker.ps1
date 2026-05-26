param(
    [string]$ExePath = "MagicMenus\bin\Debug\MagicMenus.exe",
    [int]$FrameCount = 30,
    [int]$FrameDelayMs = 30,
    [string]$Tab = "general",
    # Tab strip lives roughly at y=58..82 inside the window when on the
    # outer General tab. Override for the inner Action Menu Shell/Link strip.
    [int]$StripTop = 58,
    [int]$StripBottom = 82,
    [int]$StripLeftX = 30,
    [int]$StripRightX = 230
)

# Headlessly verify that hovering across the tab buttons no longer flashes
# white. Strategy:
#   1) Launch the app in dark-mode screenshot mode.
#   2) Find its window, bring it to the foreground.
#   3) Walk the OS cursor across the tab strip in small steps, capturing
#      a small bitmap of the tab strip after each step.
#   4) Sample pixels along a horizontal scan line through the middle of
#      the tab strip. ANY sample brighter than ~RGB(80,80,80) is treated
#      as a flicker - the dark palette caps out around RGB(60,60,65).
#   5) Report worst-case brightness and how many frames had flashes.

Get-Process MagicMenus -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 300

$env:MM_SCREENSHOT = "1"
$env:MM_SCREENSHOT_TAB = $Tab

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

Add-Type @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

public static class WinApi {
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left,Top,Right,Bottom; }
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
}
"@

$proc = Start-Process -FilePath $ExePath -PassThru
Start-Sleep -Milliseconds 1500

# Find the Magic Menu Settings window owned by our process.
$targetHwnd = [IntPtr]::Zero
$cb = [WinApi+EnumProc] {
    param($hWnd, $lParam)
    [uint32]$ownerPid = 0
    [WinApi]::GetWindowThreadProcessId($hWnd, [ref]$ownerPid) | Out-Null
    if ($ownerPid -ne $proc.Id) { return $true }
    $len = [WinApi]::GetWindowTextLength($hWnd)
    if ($len -lt 1) { return $true }
    $sb = New-Object System.Text.StringBuilder ($len + 1)
    [WinApi]::GetWindowText($hWnd, $sb, $sb.Capacity) | Out-Null
    if ($sb.ToString() -like "*Magic*") { $script:targetHwnd = $hWnd; return $false }
    return $true
}
[WinApi]::EnumWindows($cb, [IntPtr]::Zero) | Out-Null

if ($targetHwnd -eq [IntPtr]::Zero) {
    Write-Error "Could not find Magic Menus window"
    $proc.Kill()
    exit 1
}

[WinApi]::ShowWindow($targetHwnd, 5) | Out-Null
[WinApi]::BringWindowToTop($targetHwnd) | Out-Null
[WinApi]::SetForegroundWindow($targetHwnd) | Out-Null
Start-Sleep -Milliseconds 400

$r = New-Object WinApi+RECT
[WinApi]::GetWindowRect($targetHwnd, [ref]$r) | Out-Null
$winLeft = $r.Left; $winTop = $r.Top
$winW = $r.Right - $r.Left; $winH = $r.Bottom - $r.Top

$tabStripTop    = $StripTop
$tabStripBottom = $StripBottom
$tabStripLeftX  = $StripLeftX
$tabStripRightX = $StripRightX

# Capture region around just the tab strip area.
$captureLeft   = $winLeft + 10
$captureTop    = $winTop + $tabStripTop - 4
$captureWidth  = 250
$captureHeight = ($tabStripBottom - $tabStripTop) + 8

$maxBrightness = 0
$flashFrames   = 0
$results       = @()

for ($i = 0; $i -lt $FrameCount; $i++) {
    $progress = $i / [Math]::Max(1, $FrameCount - 1)
    $cursorX  = [int]($winLeft + $tabStripLeftX + ($tabStripRightX - $tabStripLeftX) * $progress)
    $cursorY  = [int]($winTop  + ($tabStripTop + $tabStripBottom) / 2)
    [WinApi]::SetCursorPos($cursorX, $cursorY) | Out-Null
    Start-Sleep -Milliseconds $FrameDelayMs

    $bmp = New-Object System.Drawing.Bitmap $captureWidth, $captureHeight
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($captureLeft, $captureTop, 0, 0, (New-Object System.Drawing.Size $captureWidth, $captureHeight))
    $g.Dispose()

    # Sample a horizontal scan line through the middle of the tab strip.
    $scanY = [int]($captureHeight / 2)
    $frameMaxBrightness = 0
    for ($x = 0; $x -lt $captureWidth; $x += 2) {
        $px = $bmp.GetPixel($x, $scanY)
        $b  = [Math]::Max([Math]::Max($px.R, $px.G), $px.B)
        if ($b -gt $frameMaxBrightness) { $frameMaxBrightness = $b }
    }
    if ($frameMaxBrightness -gt $maxBrightness) { $maxBrightness = $frameMaxBrightness }
    if ($frameMaxBrightness -gt 110) {
        $flashFrames++
        $bmp.Save((Resolve-Path .).Path + "\flicker-frame-$i.png", [System.Drawing.Imaging.ImageFormat]::Png)
    }
    $results += [PSCustomObject]@{ Frame = $i; CursorX = $cursorX; MaxBrightness = $frameMaxBrightness }
    $bmp.Dispose()
}

$proc.Kill() | Out-Null
Remove-Item Env:\MM_SCREENSHOT, Env:\MM_SCREENSHOT_TAB -ErrorAction SilentlyContinue

Write-Output ""
Write-Output ("Captured {0} frames over the tab strip, peak brightness = {1}" -f $FrameCount, $maxBrightness)
Write-Output ("Frames with white-ish flash (brightness > 110): {0}" -f $flashFrames)
if ($flashFrames -eq 0) {
    Write-Output "PASS: no flicker detected on hover."
} else {
    Write-Output "FAIL: flicker frames saved as .\flicker-frame-*.png"
}
