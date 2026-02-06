$installerPath = "$env:TEMP\VSCodeUserSetup.exe"

# Check if installer exists
if (Test-Path $installerPath) {
    Write-Host "VSCode installer found. Installing..."
    Start-Process -FilePath $installerPath -ArgumentList "/silent", "/mergetasks=!runcode" -Wait
    Write-Host "VSCode installation completed!"
} else {
    Write-Host "Installer not found. Please wait for download to complete."
}
