$Date = Get-Date -Format 'MM/dd/yyyy HH:mm:ss'
cd C:\Users\Administrator
$UserList = Get-IAMUserList
$Path = 'D:\'
$FileNameStart = 'AWS Check '
$FileNameDate = Get-Date -Format 'yyyyMMdd'
$FileExt = '.csv'
$FileName = $FileNameStart + $FileNameDate + $FileExt
$FilePath = $Path + $FileName
$ApName = "Name: ______________________________"
$ApSig = "Signature: ___________________________"
$ApDate = "Date: _______________________________"
$addOns = $ApName, $ApSig, $ApDate
$tag = '{}'

#$UserList | Export-Csv -Path $FilePath -NoTypeInformation

out-file -FilePath $FilePath -InputObject "Arn, CreateDate, PasswordLastUsed, Path, PermissionsBoundary, Tag, UserID, UserName" -Encoding Utf8

#$psObject = $null
#$psobject = New-Object psobject

foreach ($user in $UserList) {
    Out-File -FilePath $FilePath -InputObject "$($user.ARN), $($user.CreateDate), $($user.ARN), $($user.PasswordLastUsed), $($user.Path), $($user.PermissionsBoundary), $($user.UserId), $($user.UserName)" -append -Encoding Utf8
}

foreach ($add in $addOns) {
    Out-File -FilePath $FilePath -InputObject "$($add)" -append -Encoding Utf8
}

Out-File -FilePath $FilePath -InputObject "Report Generated on $date" -append -Encoding Utf8





