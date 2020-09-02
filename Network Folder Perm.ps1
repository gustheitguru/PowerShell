#Inputs 
$Date = Get-Date -Format "MM/dd/yyyy"  
$NewDate = ForEach-Object { $Date -replace "/", "." }

# Setting up your share path
#RootPath = Read-Host -Prompt "Enter Path please of Share" 

# Insert folder path where you want to save your file and its name
$OutFile = "C:\temp\Permissions-for-share-$NewDate.csv" 

#setting File Header
$Header = "Folder Path,IdentityReference,AccessControlType,IsInherited,InheritanceFlags,PropagationFlags"

#Check File Path to see if file excists
$FileExist = Test-Path $OutFile 

#If Not delete
If ($FileExist -eq $True) {Del $OutFile} 

#Setting up CSV File
Add-Content -Value $Header -Path $OutFile 

#Step through each file and write permissions to CSV 
$Folders = dir $RootPath -recurse | where {$_.psiscontainer -eq $true} foreach ($Folder in $Folders){
    $ACLs = get-acl $Folder.fullname | ForEach-Object { $_.Access  }
    Foreach ($ACL in $ACLs){
    $OutInfo = $Folder.Fullname + "," + $ACL.IdentityReference  + "," + $ACL.AccessControlType + "," + $ACL.IsInherited + "," + $ACL.InheritanceFlags + "," + $ACL.PropagationFlags
    Add-Content -Value $OutInfo -Path $OutFile 
    }} 