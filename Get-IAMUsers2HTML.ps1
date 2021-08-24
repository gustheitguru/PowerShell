$FileNameDate = Get-Date -Format 'yyyyMMdd'
$UserList = Get-IAMUserList
$FileExt = '.csv'
$PDF = '.pdf'
$html = '.html'
$Path = 'D:\CSV2PDF\'
$fileHTML = $FileNameStart + $FileNameDate + $html
$FileName = $FileNameStart + $FileNameDate + $FileExt
$Filepdf = $FileNameStart + $FileNameDate + $pdf
$FilePath = $Path + $FileName
$FilePathHTML = $Path + $FileHTML
$FilePathPDF = $Path + $filepdf
$space = '' 
$ApName = "Name:________________________________" 
$ApSig = "Signature: ___________________________"
$ApDate = "Date: _______________________________"
$RepGen = "This report was generated on " + $FileNameDate
$addOns = $space, $ApName, $ApSig, $ApDate
Write-Output $addons
$tag = '{}'

#Configurating CSV File
out-file -FilePath $FilePath -InputObject "UserName, UserID,  CreateDate, PasswordLastUsed, Arn" -Encoding Utf8


#Adding User list to CSV File
foreach ($user in $UserList) {
    Out-File -FilePath $FilePath -InputObject "$($user.UserName), $($user.UserId), $($user.CreateDate), $($user.PasswordLastUsed), $($user.arn)" -append -Encoding Utf8
}

#HTML Formatting




#convert to CSV to HTML and add HTML CSS Styling
$style = '<style>BODY{font-size: 6pt}'
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"
Import-Csv $FilePath | ConvertTo-Html -head $style | out-file $FilePathHTML

#adding in signature lines
foreach ($add in $addOns) {
    #Add-Content -Path $FilePathHTML -Value $add
    #$add1 = ConvertTo-Html -Body $add
    #Write-Output $add1
    Add-Content -Path $FilePathHTML -Value $add
    #Out-File -FilePath $FilePath -InputObject "$($add)" -append
}


######################## Convert HTML to PDF using word ###############

$wrd = new-object -com word.application 
 
# Make Word Visible 
$wrd.visible = $false
 
# Open a document  
$doc = $wrd.documents.open($FilePathHTML) 
Write-Output $FilePathHTML

# Save as pdf
$opt = 17
$name = $FilePathPDF
#Write-Output $name
$wrd.ActiveDocument.Saveas([ref]$name,[ref]$opt)

# Close and go home
$wrd.Quit()
