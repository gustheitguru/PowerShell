$Users = Get-ADUser -Filter * -Properties DisplayName, mail, telephoneNumber, department, physicalDeliveryOfficeName | select *

foreach ($User in $Users) {
    if ($User.DisplayName -like '*,*'){
     $User.name, $User.DisplayName, $user.mail, $user.telephoneNumber, $user.DistinguishedName,'---------------' | Out-File 'E:\Software_Logs\userlist.csv' -Append
    }

}
