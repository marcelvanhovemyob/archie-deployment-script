# Run first... Set-ExecutionPolicy Unrestricted
$path = $(get-location)
$xl=New-Object -ComObject Excel.Application
$wb=$xl.WorkBooks.Open("$path\file.xlsx")
$ws=$wb.WorkSheets.item(1)
$xl.Visible=$false

$timestamp = (get-date -format "HH:mm:ss")
$ws.Cells.Item(1,1)=$timestamp

$wb.SaveAs("$path\file2.xlsx")
$xl.Quit()
