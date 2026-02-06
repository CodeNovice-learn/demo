# 查找并启动微信
$paths = @(
    "C:\Program Files (x86)\Tencent\WeChat\WeChat.exe",
    "C:\Program Files\Tencent\WeChat\WeChat.exe",
    "C:\Users\zoufeng\AppData\Local\Programs\WeChat\WeChat.exe"
)

foreach ($path in $paths) {
    if (Test-Path $path) {
        Write-Host "Found WeChat at: $path"
        Start-Process $path
        Write-Host "WeChat launched!"
        exit
    }
}

Write-Host "WeChat not found. Please launch manually from Start Menu."
