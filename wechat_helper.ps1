# 查找微信安装路径
Write-Host "正在查找微信..."
$possiblePaths = @(
    "C:\Program Files (x86)\Tencent\WeChat\WeChat.exe",
    "C:\Program Files\Tencent\WeChat\WeChat.exe", 
    "C:\Users\zoufeng\AppData\Local\Programs\WeChat\WeChat.exe",
    "C:\Users\zoufeng\AppData\Local\Tencent\WeChat\WeChat.exe"
)

$wechatPath = $null

foreach ($path in $possiblePaths) {
    if (Test-Path $path) {
        $wechatPath = $path
        Write-Host "找到微信: $path"
        break
    }
}

if ($wechatPath) {
    Write-Host "正在启动微信..."
    Start-Process $wechatPath
    Write-Host "微信已启动！"
} else {
    Write-Host "未找到微信。请手动执行以下操作："
    Write-Host "1. 点击开始菜单"
    Write-Host "2. 搜索 '微信' 或 'WeChat'"
    Write-Host "3. 点击打开微信应用"
    Write-Host "4. 在微信中找到汪汪的聊天窗口"
    Write-Host "5. 点击聊天窗口右下角的 '+' 号"
    Write-Host "6. 选择 '文件' 选项"
    Write-Host "7. 浏览并选择桌面上的 '行动学习策划方案.pdf' 文件"
    Write-Host "8. 点击发送"
}
