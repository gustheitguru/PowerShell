#Files to copy
#Get Installers from folder
$APPS = (Get-ChildItem \\Gbch-dc1-vm\Shared\APPS).FullName 

New-Item -Path "c:\" -Name "apps" -ItemType "directory"

ForEach($APP in $APPS) {
	write-host $APP 
	Copy-Item $APP -Destination "c:\apps\" -Force
	
 }
