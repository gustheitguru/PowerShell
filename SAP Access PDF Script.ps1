###############################################
#                                             #
#   You need to identify what server this     #
#   will run on to merge all txt files for    #
#   the week.                                 #
#   Update file path                          #
#   setup file naming convention date         #
#                                             #
#                                             #
###############################################


$style = '<style>BODY{font-size: 6pt}'
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #C0C0C0; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "tr:nth-child(even) {background-color: #f2f2f2;}"
$style = $style + "</style>"

#add top row of text file with desired row names
@("date, ComputerName, IP, UserName") + (get-content 'D:\test\S800SAP70.txt') | Set-Content 'D:\test\user.txt'

#taking Current txt file and moving it to csv
Import-Csv -Path 'D:\test\user.txt' -Delimiter ',' | Export-Csv 'D:\test\Users.csv' -NoTypeInformation

#importing object as a varible
$data = Import-Csv 'D:\test\users.csv'
$pathHTML = 'D:\test\users.html'
$pathPDF = 'D:\test\SAP70 Access Report 20210312.pdf'


$htmlexport = @"
<html>
<head>
    $style
</head>
<body>
<h1 style="font-size:22; align-content:center">SAP70 RDP Access Report</H1>
<table border="1" style="width:100%">
<tr>
    <th>ComputerName</th>
    <th>Date</th>
    <th>ComputerName</th>
    <th>IP</th>
    <th>UserName</th>
<tr>
"@



foreach ($user in $data)
{

$htmlexport += @"
<tr>
    <td>S800SAP70</td>
    <td>$($user.Date)</td>
    <td>$($user.ComputerName)</td>
    <td>$($user.IP)</td>
    <td>$($user.UserName)</td>
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

#Send-MailMessage -From 'AWS IAM Admin User Report <NoReply_AWSIMAUsers@maruchaninc.com>' -To '<grodriguez@maruchaninc.com>' -Subject 'AWS IAM User list Report' -Body "Please review attached report for weekly IT Reports" -Attachments $pathPDF -Priority High -SmtpServer 's010net02.maruchaninc.com'

