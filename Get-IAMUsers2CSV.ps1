$Date = Get-Date -Format 'MM/dd/yyyy HH:mm:ss'
cd C:\Users\Administrator
$UserList = Get-IAMUserList
$Path = 'D:\CSV2PDF\'
$FileNameStart = 'AWS-Check-'
$FileNameDate = Get-Date -Format 'yyyyMMdd'
$FileExt = '.html'
$pdf = '.pdf'
$filePDF = $FileNameStart + $FileNameDate + $pdf
$FileName = $FileNameStart + $FileNameDate + $FileExt
$FilePath = $Path + $FileName
$space = ' ' 
$ApName = "Name: ______________________________"
$ApSig = "Signature: ___________________________"
$ApDate = "Date: _______________________________"
$addOns = $space, $ApName, $ApSig, $ApDate, $space
$tag = '{}'


#$UserList | Export-Csv -Path $FilePath -NoTypeInformation

out-file -FilePath $FilePath -InputObject "UserName, UserID,  CreateDate, PasswordLastUsed, Arn" -Encoding Utf8

#$psObject = $null
#$psobject = New-Object psobject

foreach ($user in $UserList) {
    Out-File -FilePath $FilePath -InputObject "$($user.UserName), $($user.UserId), $($user.CreateDate), $($user.PasswordLastUsed), $($user.arn)" -append -Encoding Utf8
}

foreach ($add in $addOns) {
    Out-File -FilePath $FilePath -InputObject "$($add)" -append -Encoding Utf8
}

Out-File -FilePath $FilePath -InputObject "Report Generated on $date" -append -Encoding Utf8


#CAlling Virutal excel shell 
$Exl = New-Object -ComObject Excel.Application
#launch invisible
$Exl.visible=$false

#Opening CSV in Excel Shell
$Doc = $Exl.Workbooks.Open($FilePath)


#Converting CSV2PDF
$Doc.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $filePDF)


#Closing the Document
$Doc.Close($False)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()

#Closing Excel Shell and cleaning temp Var's
$Exl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Exl)
Remove-Variable Exl


