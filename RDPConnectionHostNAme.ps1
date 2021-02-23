$Date = Get-Date
$RDPCon = Get-NetTCPConnection -LocalPort 57966 -State Established 
$RDPIP = $RDPCon.RemoteAddress
$RDPH = [System.Net.Dns]::GetHostEntry($RDPIP)
$RDPHN = $RDPH.HostName 
$Date,$RDPHN,$RDPIP -join ', ' | Out-File -FilePath 'C:\Users\bodadm\Documents\S800SAP70.txt' -Append -Width 200;

