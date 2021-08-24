#Invoke-webrequest URI Formatting. 
$URI = 'http://139.64.198.63:1337/saprdplogs'
$header = @{'Accept' = 'application/json'}
$ContentType = "application/json; charset=utf-8"

#file name configuration setup
$CSVData = 'S800SAP70_AccessLog_'
$path = 'D:\test\'
$fileNameStart = 'IPData_'
$fileDate = Get-Date -Date (get-date).AddDays(-7) -Format 'yyyyMMdd'
$filename = $fileNameStart + $filedate 
$newcsv = $path+$fileNameStart + $filedate + '.csv'
$filepath = "$path$CSVData$filedate.txt"

#merge header file with payload data for API for JSON formatting
get-content 'D:\test\fileheader.txt', $filepath| Set-Content $newcsv

#importing new CSV data file
$payload = Import-Csv $newcsv

#pushing to Strapi API DB
foreach ($item in $payload) {
    $body = $item | ConvertTo-Json

    Invoke-WebRequest -uri $URI -Method Post -Body $body -ContentType $ContentType -Headers $header
}

#deleting temp data file
Remove-Item $newcsv