#Remove all existing Powershell sessions  
Get-PSSession | Remove-PSSession  

#UserName/PWD
$UserCredential = Get-Credential

#Start Session to connecto to O365
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

#Import Session
Import-PSSession $Session

#Set Date
$Date = Get-Date -Format "MM/dd/yyyy"  
$NewDate = ForEach-Object { $Date -replace "/", "." }
#write-host $Date
write-host $NewDate

#Which Distro
$DistroName = Read-Host -Prompt "Distro Name Please..."

#The CSV Output file that is created, change for your purposes  
$OutputFile = "Distribution Group Members For $DistroName $NewDate.csv"

#Prepare Output file with headers  
Out-File -FilePath $OutputFile -InputObject "Distribution Group DisplayName,Distribution Group Email,Member DisplayName, Member Email, Member Type" -Encoding UTF8

#Get Distro Info
$DistProp = Get-DistributionGroup -Identity $DistroName
write-host  "Email = " $DistProp.PrimarySMTPAddress

#Get members of this group  
$objDGMembers = Get-DistributionGroupMember -Identity $DistroName
      



write-host "Found $($objDGMembers.Count) members..."  


#Iterate through each member  
Foreach ($objMember in $objDGMembers)  
	{  
		Out-File -FilePath $OutputFile -InputObject "$($DistProp),$($DistProp.PrimarySMTPAddress),$($objMember.DisplayName),$($objMember.PrimarySMTPAddress),$($objMember.RecipientType)" -Encoding UTF8 -append  
		write-host "`t$($DistProp),$($DistProp.PrimarySMTPAddress),$($objMember.DisplayName),$($objMember.PrimarySMTPAddress),$($objMember.RecipientType)" 
	}  
 
 
#Close up session  
Get-PSSession | Remove-PSSession  

#Resource 
#https://community.spiceworks.com/how_to/102462-office365-all-users-distribution-group

#https://gallery.technet.microsoft.com/office/List-all-Users-Distribution-7f2013b2

#https://docs.microsoft.com/en-us/powershell/module/exchange/get-dynamicdistributiongroup?view=exchange-ps

#https://www.itprotoday.com/powershell/prompting-user-input-powershell
