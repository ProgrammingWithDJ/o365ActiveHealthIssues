$AppId = ""
 $tenantId = ""
 $AppSecret = ''
    
 # Construct URI and body needed for authentication

 $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
 $body = @{
     client_id     = $AppId
     scope         = "https://graph.microsoft.com/.default"
     client_secret = $AppSecret
     grant_type    = "client_credentials" }
    
 # Get OAuth 2.0 Token

 $tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
    
 # Unpack Access Token
 $token = ($tokenRequest.Content | ConvertFrom-Json).access_token
    
 # Base URL
 $headers = @{Authorization = "Bearer $token"}
    


# Set API endpoint
$URI = "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues?$filter=isResolved eq false"

$response = (Invoke-RestMethod -Uri $URI -Headers $Headers -Method Get -ContentType "application/json") 



# Extract data from response and create a list of objects
$data = foreach ($issue in $response.value) {
    [PSCustomObject]@{
        StartDateTime = $issue.startDateTime
        EndDateTime = $issue.endDateTime
        LastModifiedDateTime = $issue.lastModifiedDateTime
        Title = $issue.title
        Id = $issue.id
        ImpactDescription = $issue.impactDescription
        Classification = $issue.classification
        Origin = $issue.origin
        Status = $issue.status
        Service = $issue.service
        Feature = $issue.feature
        FeatureGroup = $issue.featureGroup
        IsResolved = $issue.isResolved
    }
}

# Save data to CSV file
$data | Export-Csv -Path "c:\Temp\ReportsMailn.csv" -NoTypeInformation
