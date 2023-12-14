Set-ExecutionPolicy Unrestricted
# Load the Excel COM object
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open("C:\Book1")
$worksheet = $workbook.Worksheets.Item(1)

# Get the last row in the Excel sheet
$lastRow = $worksheet.UsedRange.Rows.Count
$folder = $worksheet.Cells.Item(2, 1).Value2
Add-Content C:\$folder.txt "["
Add-Content C:\$folder.txt "  {"
Add-Content C:\$folder.txt "    ""toplevel_name"": ""$folder"""
Add-Content C:\$folder.txt "  },"
# Loop through each row in the Excel sheet and assign values as variables
for ($i = 2; $i -le $lastRow; $i++) {
    $name = $worksheet.Cells.Item($i, 2).Value2
    $url = $worksheet.Cells.Item($i, 3).Value2

    # You can perform actions using these variables here
    Add-Content C:\$folder.txt "  {"
    Add-Content C:\$folder.txt "    ""url"": ""$url"","
    Add-Content C:\$folder.txt "    ""name"": ""$name"""
    Add-Content C:\$folder.txt "  },"
}
Add-Content C:\$folder.txt "]"
# Close the workbook and Excel application
$workbook.Close()
$excel.Quit()

# Release COM objects from memory
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
