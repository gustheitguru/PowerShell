#ExportTo-ExcelXPS.ps1

$path = “D:\Excel2PDF”

$xlFixedFormat = “Microsoft.Office.Interop.Excel.xlFixedFormatType” -as [type]

$excelFiles = Get-ChildItem -Path $path -include *.xls, *.xlsx -recurse

$objExcel = New-Object -ComObject excel.application

$objExcel.visible = $false

foreach($wb in $excelFiles)

    {

     $filepath = Join-Path -Path $path -ChildPath ($wb.BaseName + “.xps”)

     $workbook = $objExcel.workbooks.open($wb.fullname, 3)

     $workbook.Saved = $true

    “saving $filepath”

     $workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath)

     $objExcel.Workbooks.close()

    }

$objExcel.Quit()