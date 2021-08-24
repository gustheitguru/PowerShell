#Input File
$Input = 'D:\CSV2PDF\SAP program check 20210414.html'
#Output File
$Output = 'D:\CSV2PDF\SAP program check 20210414.pdf'

$wrd = new-object -com word.application 
 
# Make Word Visible 
$wrd.visible = $false
 
# Open a document  
$doc = $wrd.documents.open($input) 

# Save as pdf
$opt = 17
$name = $output
#Write-Output $output
$wrd.ActiveDocument.Saveas([ref]$name,[ref]$opt)

# Close and go home
$wrd.Quit()