#Invoke-webrequest URI Formatting. 
$URI = 'http://139.64.198.63:1337/Awsiamlists'
$header = @{'Accept' = 'application/json'}
$ContentType = "application/json; charset=utf-8"

#Pulling User list from AWS IAM User List
$Users = Get-IAMUserlist 

#Loop on each user collect needed Data and push to Strapi API backend
ForEach ($user in $Users){
    $groupNames = Get-IAMGroupForUser -UserName $user.UserName | select GroupName
    #$groupNames.GroupName
    $user.UserName
    $createDate = $user.CreateDate
    $cdf = $createDate.toString("MM/dd/yyyy")

    $body = ConvertFrom-StringData -StringData "UserName = $($User.UserName) `n UserID = $($User.UserId) `n CreateDate = $cdf `n PasswordLastUsed = $($user.PasswordLastUsed) `n ARN = $($user.Arn) `n Groups = $($GroupNames.GroupName)" | ConvertTo-Json
    
    $body

    Invoke-WebRequest -uri $URI -Method Post -Body $body -ContentType $ContentType -Headers $header
}


