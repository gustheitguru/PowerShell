#Get Domain
$Dom = Read-Host "Enter Domain" 

#Get Cred
$Cred =  Get-Credential

#Computer Name
$Name = Read-Host "Computer Name"

#workflow to Rename, Reboot and Add to Domain
workflow Rename-And-Reboot {
  
  param ($Name, $Dom, $Cred)
  Rename-Computer -NewName $Name -Force -Passthru
  Restart-Computer -Wait

  Add-Computer -domain $Dom –credential $Dom\$Cred -restart –force 
}