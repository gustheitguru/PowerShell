#Input File
$Input = 'D:\CSV2PDF\AWS-Check-20210416.csv'
#Output File
$Output = 'D:\CSV2PDF\AWS-Check-20210416.pdf'

#CAlling Virutal excel shell 
$Exl = New-Object -ComObject Excel.Application
#Opening CSV in Excel Shell
$Doc = $Exl.Workbooks.Open($Input)
#Converting CSV2PDF
$Doc.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $Output)
#Closing the Document
$Doc.Close($False)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()

#Closing Excel Shell and cleaning temp Var's
$Exl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Exl)
Remove-Variable Exl
Remove-Item Function:Exl-PDF