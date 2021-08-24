$style = '<style>BODY{font-size: 6pt}'
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"
$reportName = 'SAP program check '
$reportdate = get-date -Format 'MM/dd/yyyy'
$date = Get-Date 
$fileDate = Get-Date -Date ($Date).AddDays(-7) -Format 'yyyyMMdd'
$filename = 'AWS Check ' + $filedate  
$pathHTML = 'D:\csv2pdf\'+$reportName + $fileDate + '.html'
$pathPDF = 'D:\csv2pdf\' + $reportName + $fileDate + '.pdf'
Write-Output $pathHTML, $pathPDF


$htmlexport = @"
<html>
<head>
    $style
</head>
<body>
<h1 style="font-size:22; align-content:center">AWS IAM Admin User List</H1>
<table border="1" style="width:100%">
<tr>
    <th>Application Name</th>
    <th>Vendor</th>
    <th>Installed Date</th>
<tr>
"@


$applist = Get-CimInstance win32_product 


foreach ($app in $applist)
{
$htmlexport += @"
<tr>
    <td>$($app.name)</td>
    <td>$($app.Vendor)</td>
    <td>$($app.installDate)</td>
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

