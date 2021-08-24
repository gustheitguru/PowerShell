#Get-CimInstance win32_product | Select-Object Name, PackageName, InstallDate, InstallDate2 | Out-GridView
#Get-CimInstance win32_product | Get-Member -Membertype Property
#get-ciminstance win32_product | Select-Object Name, Vendor, InstallDate | Select-Object -First 1 | Out-GridView 
#get-ciminstance win32_product | Where-Object {$_.Name -eq 'Box'} | Out-GridView 

#Pulling list of installed application from local system. looking at Regedi
$Programs = Get-CimInstance win32_product 

#file Formating addons
$addOn = 'Name:________________________', 'Signature:_____________________', 'Date:______________________'
$space = '               '
$addOns = $addOn, $space, $addOn

#Date 
$Date = Get-Date 
#Actual Day M-S
$DayofWeek = ($TestDate).DayOfWeek
#acutal day represented as number 0-6
$DayofWeekNum = [int]($TestDate).DayOfWeek
#file name manipulation
$fileNameStart = 'SAP program check '
#fle name extension
$fileExtension = '.csv'
#Formating file date name
$fileDate = Get-Date -Date ($Date).AddDays(-7) -Format 'yyyyMMdd'
#concatinating name of file
$fileName = $fileNameStart + $fileDate + $fileExtension
$filePath = 'D:\test\text.csv'

Export-Csv -Path $FilePath -InputObject "Name, Vendor, InstallDate" -Encoding Utf8 -NoTypeInformation

forEach ($program in $programs) { 
    if ($Program.Vendor -like '*,*') {
        $foo = $Program.Vendor -replace ‘[,]’,”" 
    } else {
        $foo = $Program.vendor
    }

    Out-File -FilePath $FilePath -InputObject "$($program.name), $($foo), $($program.InstallDate)" -append -Encoding Utf8 
}

forEach ($Sig in $addOns) {
    Out-File -FilePath $filePath -InputObject $Sig -append -Encoding utf8 
}


