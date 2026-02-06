# Read the text file
$textFile = "C:\Users\zoufeng\.openclaw\workspace\action_learning_plan.txt"
$content = Get-Content $textFile -Raw -Encoding UTF8

# Copy to clipboard
Set-Clipboard -Value $content

Write-Host "Content copied to clipboard! Please paste it into your Word document."
