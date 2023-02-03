## Sample: MS Graph API Application Permission Test.
# Custom Values >>>
$tenantID = ""
$clientID = ""
$clientSecret = ""
$tokenURL = "https://login.microsoftonline.com/$tenantID/oauth2/token"
$resource = "https://graph.microsoft.com"
$apiVersion1 = "v1.0"
$apiVersionBeta = "beta"
# Custom Values <<<

# OAuth to get Access token >>>
# Application Permission for OAuth
$body = @{grant_type="client_credentials";resource=$resource;client_id=$clientID;client_secret=$clientSecret}
$oauth = Invoke-RestMethod -Method Post -Uri $tokenURL -Body $body
$headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
# OAuth to get Access token <<<

# MS Graph API - Get Users
$graphGetUsers = Invoke-RestMethod -Headers $headerParams -Uri "$resource/$ApiVersion1/users/"
# $graphGetUsers = Invoke-RestMethod -Headers $headerParams -Uri "$resource/$ApiVersionBeta/users/" # beta version
Write-Host "----- Get User List ------" -ForegroundColor Green
# Show Top 3.
$graphGetUsers.value[0..3] | FT -Auto
