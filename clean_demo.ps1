# 清理demo项目中的多余文件
Write-Host "正在清理demo项目..."

$demoPath = "C:\Users\zoufeng\demo"
$keepFiles = @(
    ".git",
    ".gitignore", 
    "README.md",
    "excel_reader.py",
    "requirements.txt", 
    "test_excel_reader.py"
)

Write-Host "将要保留的文件和目录:"
foreach ($file in $keepFiles) {
    $fullPath = Join-Path $demoPath $file
    if (Test-Path $fullPath) {
        Write-Host "  - $file"
    }
}

Write-Host "`n将要删除的文件:"
$items = Get-ChildItem $demoPath -Exclude $keepFiles
foreach ($item in $items) {
    if ($item.FullName -ne $demoPath) {
        Write-Host "  - $($item.Name)"
        Remove-Item $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "`n清理完成！"