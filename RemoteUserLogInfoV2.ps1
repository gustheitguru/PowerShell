
$IP = read-host -Prompt 'Input IP Address'
$User = "maruchaninc\administrator"
$PWord = ConvertTo-SecureString -String "Password1" -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord
$test = ( Get-WmiObject -Credential $Credential -Class win32_computersystem -ComputerName $ip ).Username
Write-Output $test

