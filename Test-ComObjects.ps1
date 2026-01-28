# Diagnostic script to test COM object availability
# SuperSearch Search Travis Webb V1.2026

Write-Host "Testing COM Object Availability..." -ForegroundColor Cyan
Write-Host ""

$comObjects = @(
    @{ProgId = "Word.Application"; Name = "Microsoft Word"; Required = $true},
    @{ProgId = "Excel.Application"; Name = "Microsoft Excel"; Required = $true},
    @{ProgId = "Outlook.Application"; Name = "Microsoft Outlook"; Required = $false},
    @{ProgId = "WScript.Shell"; Name = "Windows Script Host Shell"; Required = $true}
)

$allRequiredAvailable = $true

foreach ($comObj in $comObjects) {
    Write-Host "Testing $($comObj.Name) ($($comObj.ProgId))..." -NoNewline
    
    try {
        $testObj = New-Object -ComObject $comObj.ProgId -ErrorAction Stop
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testObj)
        Write-Host " [OK]" -ForegroundColor Green
    } catch [System.Management.Automation.PSArgumentException] {
        if ($comObj.Required) {
            Write-Host " [MISSING - REQUIRED]" -ForegroundColor Red
            $allRequiredAvailable = $false
        } else {
            Write-Host " [MISSING - OPTIONAL]" -ForegroundColor Yellow
        }
        Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Gray
    } catch {
        if ($comObj.Required) {
            Write-Host " [ERROR - REQUIRED]" -ForegroundColor Red
            $allRequiredAvailable = $false
        } else {
            Write-Host " [ERROR - OPTIONAL]" -ForegroundColor Yellow
        }
        Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Gray
        Write-Host "   HRESULT: $($_.Exception.HResult)" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan

if ($allRequiredAvailable) {
    Write-Host "All required COM objects are available." -ForegroundColor Green
} else {
    Write-Host "Some required COM objects are missing!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting Steps:" -ForegroundColor Yellow
    Write-Host "1. Verify Microsoft Office is installed:" -ForegroundColor White
    Write-Host "   - Check if Word and Excel are installed" -ForegroundColor Gray
    Write-Host "   - Open Word or Excel manually to verify they work" -ForegroundColor Gray
    Write-Host ""
    Write-Host "2. Repair Microsoft Office:" -ForegroundColor White
    Write-Host "   - Open Settings > Apps > Installed apps" -ForegroundColor Gray
    Write-Host "   - Find Microsoft Office" -ForegroundColor Gray
    Write-Host "   - Click the three dots > Modify" -ForegroundColor Gray
    Write-Host "   - Choose 'Quick Repair' or 'Online Repair'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "3. Re-register Office COM components (run as Administrator):" -ForegroundColor White
    Write-Host "   - Open PowerShell as Administrator" -ForegroundColor Gray
    Write-Host "   - Run: cd `"C:\Program Files\Microsoft Office\Office16`"" -ForegroundColor Gray
    Write-Host "   - Run: .\ospp.vbs /act" -ForegroundColor Gray
    Write-Host ""
    Write-Host "4. If WScript.Shell fails, repair Windows:" -ForegroundColor White
    Write-Host "   - Open Command Prompt as Administrator" -ForegroundColor Gray
    Write-Host "   - Run: sfc /scannow" -ForegroundColor Gray
    Write-Host "   - Run: DISM /Online /Cleanup-Image /RestoreHealth" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
