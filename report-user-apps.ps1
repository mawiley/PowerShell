#Variables to configure
$tenantID = "" #your tenantID or tenant root domain
$appID = "" #the GUID of your app. For best result, use app with Team.ReadBasic.All, TeamsAppInstallation.ReadForTeam.All, TeamsAppInstallation.ReadForUser.All and TeamsTab.Read.All scopes granted.
$client_secret = "" #client secret for the app

#Prepare token request
$url = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'

$body = @{
    grant_type = "client_credentials"
    client_id = $appID
    client_secret = $client_secret
    scope = "https://graph.microsoft.com/.default"
}

#Obtain the token 
try { $tokenRequest = Invoke-WebRequest -Method Post -Uri $url -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -ErrorAction Stop }
catch { Write-Host "Unable to obtain access token, aborting..."; return }

$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

$authHeader1 = @{
   'Content-Type'='application\json'
   'Authorization'="Bearer $token"
}

#Get a list of all Users
$Users = @()
$uri = "https://graph.microsoft.com/beta/users?`$select=id,displayName"
do {
    $result = Invoke-WebRequest -Headers $AuthHeader1 -Uri $uri -ErrorAction Stop
    $uri = $result.'@odata.nextLink'
    #If we are getting multiple pages, best add some delay to avoid throttling
    Start-Sleep -Milliseconds 500
    $Users += ($result.Content | ConvertFrom-Json).Value
} while ($uri)
if (!$Users -or $Users.Count -eq 0) { Write-Host "Unable to obtain the list of users, exiting..."; return }

#Iterate over each Team and prepare the report
$ReportApps = @();
$count = 1
foreach ($user in $Users) {

    #Progress message
    $ActivityMessage = "Retrieving data for user $($user.displayName). Please wait..."
    $StatusMessage = ("Processing user {0} of {1}: {2}" -f $count, @($Users).count, $user.id)
    $PercentComplete = ($count / @($Users).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Simple anti-throttling control
    Start-Sleep -Milliseconds 500

    #get a list of apps for the Team
    $userApps = Invoke-WebRequest -Headers $authHeader1 -Uri "https://graph.microsoft.com/beta/users/$($user.id)/teamwork/installedApps?`$expand=teamsApp" -ErrorAction Stop
    $userApps = ($userApps.Content | ConvertFrom-Json).Value

    $i = 0
    foreach ($app in $userApps) {
        $i++
        $objApp = New-Object PSObject
        $objApp | Add-Member -MemberType NoteProperty -Name "Number" -Value $i
        $objApp | Add-Member -MemberType NoteProperty -Name "User" -Value $user.displayName
        $objApp | Add-Member -MemberType NoteProperty -Name "AppName" -Value $app.teamsApp.displayName
        $objApp | Add-Member -MemberType NoteProperty -Name "AppId" -Value $app.teamsApp.id
        $objApp | Add-Member -MemberType NoteProperty -Name "AddedVia" -Value $app.teamsApp.distributionMethod

        $ReportApps += $objApp
    }
}

#Export the result
$ReportApps | Select-Object * -ExcludeProperty Number | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_UsersAppsReport.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
