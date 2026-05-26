param(
    [string]$ExePath = "MagicMenus\bin\Debug\MagicMenus.exe",
    [string]$Prefix = "shot",
    [int]$WaitMs = 2500,
    [string[]]$Tabs = @("general", "general-settings", "action", "action-link", "clipboard", "clipboard-string")
)

# Capture each requested view, one at a time, with the redesigned UI.

foreach ($tab in $Tabs) {
    Get-Process MagicMenus -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 300
    $env:MM_SCREENSHOT = "1"
    $env:MM_SCREENSHOT_TAB = $tab
    & powershell -ExecutionPolicy Bypass -File .\capture-screenshot.ps1 -OutPath ".\$Prefix-$tab.png" -WaitMs $WaitMs
}

Remove-Item Env:\MM_SCREENSHOT -ErrorAction SilentlyContinue
Remove-Item Env:\MM_SCREENSHOT_TAB -ErrorAction SilentlyContinue
