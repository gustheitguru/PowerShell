$row1 = "Name: _______________________________  Sig: _______________________________  Date: _______________________________"
$row2 = "Name: _______________________________  Sig: _______________________________  Date: _______________________________"
$rows = $row1, $row2

$Path = 'D:\'
$FileNameStart = 'Domain Admin Review '
$FileNameDate = Get-Date -Format 'yyyyMMdd'
$FileExt = '.csv'
$FileName = $FileNameStart + $FileNameDate + $FileExt
$FilePath = $Path + $FileName
$UserList = Get-ADGroupMember -Identity 'Domain Admins'

out-file -FilePath $FilePath -InputObject "Name, UserName, ObjectGUID" -Encoding Utf8

foreach ($user in $UserList) {
    Out-File -FilePath $FilePath -InputObject "$($user.name), $($user.SamAccountName), $($user.ObjectGUID)" -append -Encoding Utf8
}

foreach ($row in $rows) {
    Out-File -FilePath $FilePath -InputObject "$($row)" -append -Encoding Utf8
}

Out-File -FilePath $FilePath -InputObject "Report Generated on $date" -append -Encoding Utf8