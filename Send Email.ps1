###################################
#                                 #
#       Sending Mail Script       #
# From S800SAP70 to email Access  #
#  Log Report RDP Connections     #
#                                 # 
#                                 #
###################################

# Flag Break Down
# From - email from system
# To - Who will receive email 
# Subject - Email Subject
# BOdy - Email Body
# Attachemnet - Attached Report from location
# SMTP - SMTP Server
# Priority - Email Priority


Send-MailMessage -From 'Access Report S800SAP70 <NoReply_AccessReportSAP70@maruchaninc.com>' -To 'Gus <grodriguez@maruchaninc.com>' -Subject 'Access Log Report S800SAP70' -Body "Please review attached report for weekly IT Reports" -Attachments 'D:\RDPLog\test.csv' -Priority High -SmtpServer 's010net02.maruchaninc.com'

