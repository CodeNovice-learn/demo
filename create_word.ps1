$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $true
$doc = $wordApp.Documents.Add()
$doc.SaveAs([environment]::GetFolderPath('Desktop') + '\新建文档.docx')
