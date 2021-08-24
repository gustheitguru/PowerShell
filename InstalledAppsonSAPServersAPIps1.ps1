$apps = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, InstallDate, Publisher 
# Get-WmiObject Win32_Product -ComputerName 172.22.24.4 | Select-Object -Property Name, Vendor, InstallDate
#Invoke-webrequest URI Formatting. 
$URI = 'http://139.64.198.63:1337/sias'
$header = @{'Accept' = 'application/json'}
$ContentType = "application/json; charset=utf-8"



foreach ($app in $apps) {
   
   if($app.InstallDate) {
        $body = ConvertFrom-StringData -StringData "DisplayName = $($app.DisplayName) `n DisplayVersion = $($app.DisplayVersion) `n Publisher = $($app.Publisher) `n InstallDate = $($app.InstallDate) `n servernamea = TestServer2" | ConvertTo-Json
        $body
        Invoke-WebRequest -uri $URI -Method Post -Body $body -ContentType $ContentType -Headers $header
    }


}


