$desktopPath = [Environment]::GetFolderPath('Desktop')
$installerPath = Join-Path $desktopPath "VSCodeUserSetup.exe"

Write-Host "Downloading VSCode to desktop..."
Invoke-WebRequest -Uri 'https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user' -OutFile $installerPath

Write-Host "Download complete! Starting installation..."
Start-Process -FilePath $installerPath -Wait

Write-Host "VSCode installation completed!"
