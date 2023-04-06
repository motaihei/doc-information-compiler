# Get running Excel application object
$Excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')

# Get the addresses of all open Excel workbooks with .xlsx or .xls extension
$Addresses = @($Excel.Workbooks | Where-Object {$_.FullName -like '*.xls*' -and ($_.FullName -like '*.xlsx' -or $_.FullName -like '*.xls') } | ForEach-Object {$_.FullName})

# Output the addresses of all open Excel workbooks
Write-Output $Addresses