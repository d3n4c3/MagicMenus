param(
    [string]$ExePath = "MagicMenus\bin\Debug\MagicMenus.exe",
    [string]$OutPath = "screenshot.png",
    [int]$WaitMs = 3000
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public static class WinApi {
    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd, int nCmdShow);
    [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetWindowText(IntPtr hwnd, System.Text.StringBuilder str, int max);
    [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr hwnd);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    public static List<IntPtr> GetWindowsForProcess(uint pid) {
        var result = new List<IntPtr>();
        EnumWindows(delegate(IntPtr hwnd, IntPtr lParam) {
            uint wpid;
            GetWindowThreadProcessId(hwnd, out wpid);
            if (wpid == pid && IsWindowVisible(hwnd)) {
                RECT r;
                if (GetWindowRect(hwnd, out r) && (r.Right - r.Left) > 50 && (r.Bottom - r.Top) > 50) {
                    result.Add(hwnd);
                }
            }
            return true;
        }, IntPtr.Zero);
        return result;
    }
}
"@

$env:MM_SCREENSHOT = "1"
# Allow MM_SCREENSHOT_TAB to be passed through from the calling scope.
$proc = Start-Process -FilePath $ExePath -PassThru
Start-Sleep -Milliseconds $WaitMs

try {
    $hwnd = [IntPtr]::Zero
    for ($i = 0; $i -lt 30; $i++) {
        $windows = [WinApi]::GetWindowsForProcess([uint32]$proc.Id)
        if ($windows.Count -gt 0) {
            # Prefer the largest visible top-level window for this process.
            $bestArea = 0
            foreach ($h in $windows) {
                $r = New-Object WinApi+RECT
                [WinApi]::GetWindowRect($h, [ref]$r) | Out-Null
                $area = ($r.Right - $r.Left) * ($r.Bottom - $r.Top)
                if ($area -gt $bestArea) { $bestArea = $area; $hwnd = $h }
            }
            if ($hwnd -ne [IntPtr]::Zero) { break }
        }
        Start-Sleep -Milliseconds 200
    }

    if ($hwnd -eq [IntPtr]::Zero) {
        Write-Output "Could not find a visible MagicMenus window."
        return
    }

    [WinApi]::ShowWindow($hwnd, 9) | Out-Null   # SW_RESTORE
    [WinApi]::SetForegroundWindow($hwnd) | Out-Null
    [WinApi]::BringWindowToTop($hwnd) | Out-Null
    Start-Sleep -Milliseconds 600

    $rect = New-Object WinApi+RECT
    [WinApi]::GetWindowRect($hwnd, [ref]$rect) | Out-Null

    # Pad so we can see the surrounding area / drop shadow.
    $pad = 20
    $left = $rect.Left - $pad
    $top = $rect.Top - $pad
    $w = ($rect.Right - $rect.Left) + 2 * $pad
    $h = ($rect.Bottom - $rect.Top) + 2 * $pad

    $screen = [System.Windows.Forms.Screen]::FromHandle($hwnd).Bounds
    if ($left -lt $screen.Left) { $w -= ($screen.Left - $left); $left = $screen.Left }
    if ($top  -lt $screen.Top)  { $h -= ($screen.Top  - $top);  $top  = $screen.Top  }
    if ($left + $w -gt $screen.Right)  { $w = $screen.Right  - $left }
    if ($top  + $h -gt $screen.Bottom) { $h = $screen.Bottom - $top  }

    $bmp = New-Object System.Drawing.Bitmap $w, $h
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($left, $top, 0, 0, (New-Object System.Drawing.Size $w, $h))
    $g.Dispose()
    $bmp.Save($OutPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Output ("Captured hwnd 0x{0:X} at ({1},{2}) size {3}x{4} -> {5}" -f $hwnd.ToInt64(), $left, $top, $w, $h, $OutPath)
}
finally {
    if (-not $proc.HasExited) {
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    }
    Remove-Item Env:\MM_SCREENSHOT -ErrorAction SilentlyContinue
}
