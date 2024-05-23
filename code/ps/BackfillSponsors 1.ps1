#Import-Module Microsoft.Graph.Beta
#Import-Module MSAL.PS

# Useful link https://tech.nicolonsky.ch/explaining-microsoft-graph-access-token-acquisition/
# For app configuration https://morgantechspace.com/2022/03/azure-ad-get-access-token-for-delegated-permissions-using-powershell.html
# For pagination https://medium.com/@mozzeph/how-to-handle-microsoft-graph-paging-in-powershell-354663d4b32a

$useDefault= Read-Host "Use default values for app id and tenant id (Y/N)?"
if($useDefault -eq 'Y')
{
    # Old appId 7280d3bf-8456-46db-91bb-f3df2d6df5f0
    # Currently using Microsoft Graph PowerShell app id
    $appId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
    $tenantId = "926d99e7-117c-4a6a-8031-0cc481e9da26"
}
else
{
    $appId = Read-Host "What is the app id to use for getting token?"
    $tenantId = Read-Host "What is the tenant id?"
}

$whatIf= Read-Host "Use WhatIf mode (Y/N)?"
if($whatIf -eq 'Y')
{
    $whatIf = $true
}
else
{
    $whatIf = $false
}
Connect-MgGraph -TenantId $tenantId -Scopes User.ReadWrite.All
# $context = Get-MgContext 

function Get-InvitedBy([string]$userId,$authHeader){

    $invitedByUrl = 'https://graph.microsoft.com/beta/users/' + $userId + '/invitedBy' 
    $response = Invoke-RestMethod -Uri $invitedByUrl -Headers $authHeader
    if($response.value.Count -gt 1)
    {
        Write-Error "Error: More than 1 invitedBy found, not updating for userId:" $userId
        return $null
    }

    # return the object id
    return $response.value.Id

}

# Get external users 
$guestUsers = Get-MgBetaUser -Filter "creationType eq 'Invitation'" -Property Id,UserPrincipalName -ExpandProperty Sponsors -Top 5

# Get token to call invitedBy api
$connectionDetails = @{
    'TenantId'    = $tenantId
    'ClientId'    = $appId
    'Interactive' = $true
}
    
$token = Get-MsalToken @connectionDetails -Scopes User.Read.All
$authHeader = @{'Authorization' = $token.CreateAuthorizationHeader()}
$backfillCount = 0

# Get invitedBy for each of the external user and use it to populate sponsors
foreach ($guestUser in $guestUsers)
{
    $userId = $guestUser.Id
    # If user already has sponsors do not update using invitedBy
    if($guestUser.Sponsors.Count -ge 1)
    {
        Write-Host "SponsorsPresent: Not adding sponsors for userId:" $userId
        continue
    }

    $invitedById = Get-InvitedBy -userId $userId -authHeader $authHeader
    if($null -eq $invitedById)
    {
        Write-Host "Error: InvitedBy not populated for user:" $userId
        continue
    }

    $sponsorUrl = "https://graph.microsoft.com/beta/users/" + $invitedById 
    $dirObj = @{"sponsors@odata.bind" = @($sponsorUrl) }
    $sponsorsRequestBody = $dirObj | ConvertTo-Json

    # Backfill user's sponsors with invitedBy data
    if($whatIf)
    {
        Write-Host "WhatIf: Would have updated sponsors with invitedBy id:" $invitedById  " userId:" $userId
        $backfillCount++   
        continue
    }
    else
    {
        Update-MgBetaUser -UserId $userId  -BodyParameter $sponsorsRequestBody
        Write-Host "Success: Updated sponsors with invitedBy id:" $invitedById  " userId:" $userId
        $backfillCount++
    }
}

Write-Host "Success: Backfill complete. WhatIf mode:" $whatIf "backfill count: " $backfillCount

