#Date 
$Date = Get-Date 
#Actual Day M-S
$DayofWeek = ($TestDate).DayOfWeek
#acutal day represented as number 0-6
$DayofWeekNum = [int]($TestDate).DayOfWeek
#file name manipulation
$fileNameStart = 'S800SAP70_AccessLog_'
#fle name extension
$fileExtension = '.csv'
$fileExtPDF = '.pdf'
#file name date verification
$DayofWeekNumAdj = if($DayofWeekNum -eq '2' -or $DayofWeekNum -eq '3' -or $DayofWeekNum -eq '4' -or $DayofWeekNum -eq '5' -or $DayofWeekNum -eq '6' ){ (($DayofWeekNum - 1) * -1) } elseif ($DayofWeekNum -eq 0 ) { '-6' } else { '0' }
#Formating file date name
$fileDate = Get-Date -Date ($TestDate).AddDays($DayofWeekNumAdj) -Format 'yyyyMMdd'
#concatinating name of file
$fileName = $fileNameStart + $fileDate + $fileExtension
$fileNamePDF = $fileNameStart + $fileDate + $fileExtPDF
#set file path
$filePath = 'D:\RDPLog\'+$fileName
$filePathDPF = 'D:\RDPLog\'+$fileNamePDF

if (Test-Path -Path $filePath -PathType Leaf) {
    #Convert csv to PDF and email PDF

    #Email new PDF
    Send-MailMessage -From 'NoReply_RDPLog@maruchaninc.com' -To 'grodriguez@maruchaninc.com' -Subject 'Weekly RDP Log Report S800SAP70' -Body 'This is the weekly Report for RDP activity on S800SAP70' -Attachments $filePathDPF -SmtpServer 's010net02.maruchaninc.com'
     
    } else { 
    #email there is an issue with logging
     Send-MailMessage -From 'NoReply_RDPLog@maruchaninc.com' -To 'grodriguez@maruchaninc.com' -Subject 'No Weekly RDP Log Report S800SAP70' -Body 'There is no RDP report for this week. Please double check server Event logs' -SmtpServer 's010net02.maruchaninc.com'
    }




# File paths
$txtPath = $filePath
$pdfPath = $fileName = $fileNameStart + $fileDate + '.pdf'

# Required Word Variables
$wdExportFormatPDF = 17
$wdDoNotSaveChanges = 0

#add A excel book
$objExcel = New-Object -ComObject excel.application
$objExcel.visible = $false
$workbook = $filePath
Write-Output $workbook

# Export the PDF file and close without saving a Word document
$workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, 'D:\RDPLog')
$workbook.close([ref]$wdDoNotSaveChanges)
$objExcel.Quit()