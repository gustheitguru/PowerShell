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

#Get all Distribution Groups from Office 365  
$objDistributionGroups = Get-DistributionGroup -ResultSize Unlimited  
  
#Iterate through all groups, one at a time      
Foreach ($objDistributionGroup in $objDistributionGroups)  
{      
     
    write-host "Processing $($objDistributionGroup.DisplayName)..."  
  
    #Get members of this group  
    $objDGMembers = Get-DistributionGroupMember -Identity $($objDistributionGroup.PrimarySmtpAddress)  
      
    write-host "Found $($objDGMembers.Count) members..."  
      
    #Iterate through each member  
    Foreach ($objMember in $objDGMembers)  
    {  
        Out-File -FilePath $OutputFile -InputObject "$($objDistributionGroup.DisplayName),$($objDistributionGroup.PrimarySMTPAddress),$($objMember.DisplayName),$($objMember.PrimarySMTPAddress),$($objMember.RecipientType)" -Encoding UTF8 -append  
        write-host "`t$($objDistributionGroup.DisplayName),$($objDistributionGroup.PrimarySMTPAddress),$($objMember.DisplayName),$($objMember.PrimarySMTPAddress),$($objMember.RecipientType)" 
    }  
}  



#Close up session  
Get-PSSession | Remove-PSSession 