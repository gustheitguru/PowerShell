
#remove files for Global connect

#pull current logged on user
$S = (gwmi win32_loggedonuser).antecedent.split('=')
$name = $s[2].Replace('"','') ## username

#settup path with current logged on username
$path = (-join('C:\Users\', "$name",'\AppData\Local\Palo Alto Networks\GlobalProtect\*' ))

#remove the item
remove-item -Path $path -Exclude "*.log"

