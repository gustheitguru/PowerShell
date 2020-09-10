#Install Remote Packages
#Chrome,  MSOffice, eCare, LTC, AHT

#Souce Location \\GBCH-DC1-VM\SHARED\APPS
#Destination C:\Install

$dest = "C:\Install"



#Get Installers from folder
$APPS = Get-ChildItem \\Gbch-dc1-vm\shared\APPS -Name | ForEach-Object {
 write-host $_
 start-process $_
 }

