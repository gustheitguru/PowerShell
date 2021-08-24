$style = '<style>BODY{font-size: 6pt}'
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #C0C0C0; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "tr:nth-child(even) {background-color: #f2f2f2;}"
$style = $style + "</style>"
$reportName = 'AWS Check '
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
    <th>UserName</th>
    <th>UserId</th>
    <th>CreateDate</th>
    <th>PasswordLastUsed</th>
    <th>ARN</th>
    <th>Groups</th>
<tr>
"@


$User = Get-IAMUserList


foreach ($user in $users)
{
    $groupNames = Get-IAMGroupForUser -UserName $user.UserName | select GroupName

$htmlexport += @"
<tr>
    <td>$($user.UserName)</td>
    <td>$($user.UserId)</td>
    <td>$($user.CreateDate)</td>
    <td>$($user.PasswordLastUsed)</td>
    <td>$($user.arn)</td>
    <td>$($groupNames.GroupName)</td>
</tr>
"@
}
$htmlexport += @"
</table>
<p style="font-size:12">Name:__________________________________</p>
<p style="font-size:12">Signature:_______________________________</p>
<p style="font-size:12">Date:___________________________________</p>
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

