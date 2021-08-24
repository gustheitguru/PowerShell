#Get-ADGroupMember -Identity 'Domain Admins'
$load = 'D:\test\dump.txt'
$payload = Import-Excel -Path $load
$payload 
