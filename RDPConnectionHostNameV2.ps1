#Date 
$Date = Get-Date 
#Actual Day M-S
$DayofWeek = ($TestDate).DayOfWeek
#acutal day represented as number 0-6
$DayofWeekNum = [int]($TestDate).DayOfWeek

#Connecting User information
#$RDPCon = Get-NetTCPConnection -LocalPort 57966 -State Established 
#$RDPIP = $RDPCon.RemoteAddress
#$RDPH = [System.Net.Dns]::GetHostEntry($RDPIP)
#$RDPHN = $RDPH.HostName 

#file name manipulation
$fileNameStart = 'S800SAP70_AccessLog_'
#fle name extension
$fileExtension = '.csv'
#file name date verification
$DayofWeekNumAdj = if($DayofWeekNum -eq '2' -or $DayofWeekNum -eq '3' -or $DayofWeekNum -eq '4' -or $DayofWeekNum -eq '5' -or $DayofWeekNum -eq '6' ){ (($DayofWeekNum - 1) * -1) } elseif ($DayofWeekNum -eq 0 ) { '-6' } else { '0' }
#Formating file date name
$fileDate = Get-Date -Date ($TestDate).AddDays($DayofWeekNumAdj) -Format 'yyyyMMdd'
#concatinating name of file
$fileName = $fileNameStart + $fileDate + $fileExtension
#set file path
$filePath = 'D:\RDPLog\'+$fileName

#writing log information to file
$Date,$RDPHN,$RDPIP -join ', ' | Out-File -FilePath $filePath -Append -Width 200;

Write-Output $filePath
Write-Output $Date



