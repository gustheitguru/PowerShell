#enter Message
$Message = Read-Host Enter Message

#Send Messgae to computers on Domain and message will last 15 minutes on the screen. 
#You can remove the /TIME: to leave message on screen for ever

(Get-ADComputer -SearchBase "OU=Computers,OU=MyBusiness,DC=gbchdomain,DC=local" -Filter *).Name | Foreach-Object {
	write-host ----------------------------
	msg * /server:$_ /TIME:900 "$Message"
	write-host $_
	write-host ____________________________
}