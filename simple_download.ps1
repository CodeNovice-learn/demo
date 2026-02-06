# Simple download script
$desktopPath = [Environment]::GetFolderPath('Desktop')
Write-Host "Desktop path: $desktopPath"

# Check if VSCode installer exists on desktop
$desktopInstaller = Join-Path $desktopPath "VSCodeUserSetup.exe"
if (Test-Path $desktopInstaller) {
    Write-Host "VSCode installer found on desktop. Starting installation..."
    Start-Process -FilePath $desktopInstaller -ArgumentList "/silent", "/mergetasks=!runcode" -Wait
    Write-Host "VSCode installation completed!"
} else {
    Write-Host "VSCode installer not found. Please download manually from https://code.visualstudio.com/download"
}
