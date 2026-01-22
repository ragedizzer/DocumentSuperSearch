#SuperSearch Search Travis Webb V1.2026
#Create a desktop shortcut for the Search GUI launcher

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$guiPath = Join-Path -Path $scriptDir -ChildPath "Search-Gui.ps1"
if (-not (Test-Path -LiteralPath $guiPath)) {
    Write-Error "Search-Gui.ps1 not found in $scriptDir"
    return
}

$desktopPath = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrWhiteSpace($desktopPath)) {
    Write-Error "Desktop folder not found for this user."
    return
}

$shortcutPath = Join-Path -Path $desktopPath -ChildPath "Search GUI.lnk"

$shell = $null
$shortcut = $null
try {
    $shell = New-Object -ComObject WScript.Shell
$powerShellPath = Join-Path -Path $env:SystemRoot -ChildPath "System32\WindowsPowerShell\v1.0\powershell.exe"
if (-not (Test-Path -LiteralPath $powerShellPath)) {
    $powerShellPath = "powershell.exe"
}

$arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$guiPath`""

$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $powerShellPath
$shortcut.Arguments = $arguments
    $shortcut.WorkingDirectory = $scriptDir
$shortcut.IconLocation = "$powerShellPath,0"
    $shortcut.Save()
    Write-Host "Shortcut created: $shortcutPath" -ForegroundColor Green
} catch {
    Write-Error "Failed to create shortcut: $($_.Exception.Message)"
} finally {
    if ($shortcut -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shortcut) | Out-Null
    }
    if ($shell -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
    }
}