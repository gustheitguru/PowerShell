$items = Get-ChildItem -Path D:\IPCheck

foreach($item in $items){
    $XLSString = Select-String -Path $item -Pattern '139.64.200.15'
    $ACCString = Select-String -Path $item -Encoding unicode -Pattern '139.64.200.15'
    $extnXLS = [IO.PATH]::GetExtension($item ) -eq '.xls'
    $extnACC = [IO.PATH]::GetExtension($item ) -eq '.accdb'

    if ($XLSString -and $extnXLS) {

        $item | Select-Object -Property Name | Export-Csv -Path D:\IPCheck\ExcelWithIP.csv -NoTypeInformation -Append
        
    }
    elseif ($ACCString -and $extnACC) {
        $item | Select-Object -Property Name | Export-Csv -Path D:\IPCheck\AccessDBWithIP.csv -NoTypeInformation -Append
    }

}