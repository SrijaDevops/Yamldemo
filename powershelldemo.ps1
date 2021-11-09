$path = "C:\Users\User\Documents\exps.xlsx"

$Excel = New-Object -Com Excel.Application
$Excel.Visible = $true
$Workbook = $Excel.Workbooks.Open($Path)
$page = 'Sheet2'
$ws = $Workbook.worksheets | where-object {$_.Name -eq $page}