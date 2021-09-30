# open Windows Dialog box to accept a new file

Add-Type -AssemblyName System.Windows.Forms

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }

$null = $FileBrowser.ShowDialog()

# $FileBrowser - will show details of file
# $FileBrowser.FileName - will show the Path of the file