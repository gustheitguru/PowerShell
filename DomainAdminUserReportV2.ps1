#Date 
$Date = Get-Date 
#Actual Day M-S
$DayofWeek = ($Date).DayOfWeek
#acutal day represented as number 0-6
$DayofWeekNum = [int]($Date).DayOfWeek
#file name manipulation
$ReportName = 'process '
#fle name extension
$fileExtension = '.csv'
$fileExtPDF = '.pdf'
#file name date verification
$DayofWeekNumAdj = if($DayofWeekNum -eq '2' -or $DayofWeekNum -eq '3' -or $DayofWeekNum -eq '4' -or $DayofWeekNum -eq '5' -or $DayofWeekNum -eq '6' ){ (($DayofWeekNum - 1) * -1) } elseif ($DayofWeekNum -eq 0 ) { '-6' } else { '0' }
#Formating file date name
$fileDate = Get-Date -Date ($Date).AddDays($DayofWeekNumAdj) -Format 'yyyyMMdd'
#file Path
$pathHTML = 'D:\csv2pdf\'+$reportName + $fileDate + '.html'
$pathPDF = 'D:\csv2pdf\' + $reportName + $fileDate + '.pdf'
Write-Output $pathHTML, $pathPDF


#CSS Formatting for HTML File 
$style = '<style>BODY{font-size: 6pt}'
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"


#sending over HTML body to file
$htmlexport = @"
<html>
<head>
    $style
</head>
<body>
<h1 style="font-size:22; align-content:center">Domain Admin User List</H1>
<table border="1" style="width:100%">
<tr>
    <th>Account Name</th>
    <th>UserId</th>
    <th>ObjectClass</th>
    <th>distinguishedName</th>
<tr>
"@


$Users = Get-process | Select-Object -First 5

foreach ($user in $users)
{
$htmlexport += @"
<tr>
    <td>$($user.handel)</td>
    <td>$($user.id)</td>
    <td>$($user.ProcessName)</td>
    <td>$($user.si)</td>
</tr>
"@
}
$htmlexport += @"
</table>
<p style="font-size:12">Name:________________________________    Signature:_____________________________  Date:__________________</p>
<p style="font-size:12">Name:________________________________    Signature:_____________________________  Date:__________________</p>
<p style="font-size:12">This report was generated on $($reportdate)</p>
</body>
</html>
"@

$htmlexport | Out-File $pathHTML

<#

######################## Convert HTML to PDF using word ###############

$wrd = new-object -com word.application 
 
# Make Word Visible 
$wrd.visible = $false
 
# Open a document  
$doc = $wrd.documents.open($pathHTML) 

# Save as pdf
$opt = 17
$name = $pathPDF
#Write-Output $name
$wrd.ActiveDocument.Saveas([ref]$name,[ref]$opt)

# Close and go home
$wrd.Quit()


###################### Send PDF as Email ############################
#                                                                   #
# Flag Break Down                                                   #
# From - email from system                                          #
# To - Who will receive email                                       #
# Subject - Email Subject                                           #
# BOdy - Email Body                                                 #
# Attachemnet - Attached Report from location                       #
# SMTP - SMTP Server                                                #
# Priority - Email Priority                                         #
#                                                                   #
#####################################################################

Send-MailMessage -From 'AWS IAM Admin User Report <NoReply_AWSIMAUsers@maruchaninc.com>' -To '<grodriguez@maruchaninc.com>' -Subject 'AWS IAM User list Report' -Body "Please review attached report for weekly IT Reports" -Attachments $pathPDF -Priority High -SmtpServer 's010net02.maruchaninc.com'

#>