$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -ge 7) {
    $PSNativeCommandUseErrorActionPreference = $true
}

function Invoke-GraphJson {
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $false)][object]$Body
    )

    if ($null -eq $Body) {
        return az rest --method $Method --url $Url
    }

    $tmp = [IO.Path]::GetTempFileName()
    try {
        $json = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 20 }
        Set-Content -Path $tmp -Value $json -NoNewline -Encoding UTF8
        return az rest --method $Method --url $Url --headers 'Content-Type=application/json' --body "@$tmp"
    } finally {
        Remove-Item $tmp -ErrorAction SilentlyContinue
    }
}

$displayName = 'Agent 365 CLI'
$graphAppId = '00000003-0000-0000-c000-000000000000'
$scopeNames = @(
    'AgentIdentityBlueprintPrincipal.Create',
    'AgentIdentityBlueprint.ReadWrite.All',
    'AgentIdentityBlueprint.UpdateAuthProperties.All',
    'AgentIdentityBlueprint.AddRemoveCreds.All',
    'AgentIdentityBlueprint.DeleteRestore.All',
    'AgentInstance.ReadWrite.All',
    'AgentIdentity.DeleteRestore.All',
    'DelegatedPermissionGrant.ReadWrite.All',
    'Directory.Read.All',
    'User.Read'
)

$existing = az ad app list --display-name $displayName --query '[0]' -o json | ConvertFrom-Json

$graphSp = (az rest --method GET --url "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$graphAppId'&`$select=id,oauth2PermissionScopes" | ConvertFrom-Json).value[0]
$resourceAccess = @()
foreach ($name in $scopeNames) {
    $scope = $graphSp.oauth2PermissionScopes | Where-Object { $_.value -eq $name }
    if (-not $scope) {
        throw "Missing Microsoft Graph delegated scope metadata: $name"
    }
    $resourceAccess += @{ id = $scope.id; type = 'Scope' }
}

if ($existing) {
    $appId = $existing.appId
    $appObjectId = $existing.id
    Write-Host "Found existing app registration: $appId"
} else {
    $body = @{
        displayName = $displayName
        signInAudience = 'AzureADMyOrg'
        publicClient = @{ redirectUris = @('http://localhost:8400/', 'http://localhost') }
        isFallbackPublicClient = $true
        requiredResourceAccess = @(@{
            resourceAppId = $graphAppId
            resourceAccess = $resourceAccess
        })
    } | ConvertTo-Json -Depth 20

    $created = Invoke-GraphJson -Method POST -Url 'https://graph.microsoft.com/v1.0/applications' -Body $body | ConvertFrom-Json
    $appId = $created.appId
    $appObjectId = $created.id
    Write-Host "Created app registration: $appId"
}

$patch = @{
    publicClient = @{
        redirectUris = @(
            'http://localhost:8400/',
            'http://localhost',
            "ms-appx-web://Microsoft.AAD.BrokerPlugin/$appId"
        )
    }
    isFallbackPublicClient = $true
    requiredResourceAccess = @(@{
        resourceAppId = $graphAppId
        resourceAccess = $resourceAccess
    })
} | ConvertTo-Json -Depth 20
Invoke-GraphJson -Method PATCH -Url "https://graph.microsoft.com/v1.0/applications/$appObjectId" -Body $patch | Out-Null

$clientSp = (az rest --method GET --url "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$appId'&`$select=id,appId,displayName" | ConvertFrom-Json).value[0]
if (-not $clientSp) {
    $clientSp = Invoke-GraphJson -Method POST -Url 'https://graph.microsoft.com/v1.0/servicePrincipals' -Body @{ appId = $appId } | ConvertFrom-Json
    Write-Host "Created service principal: $($clientSp.id)"
}

$scopeString = ($scopeNames -join ' ')
$grant = (az rest --method GET --url "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=clientId eq '$($clientSp.id)' and resourceId eq '$($graphSp.id)'" | ConvertFrom-Json).value[0]
if ($grant) {
    Invoke-GraphJson -Method PATCH -Url "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$($grant.id)" -Body @{ scope = $scopeString } | Out-Null
    Write-Host "Updated delegated permission grant: $($grant.id)"
} else {
    $grantBody = @{
        clientId = $clientSp.id
        consentType = 'AllPrincipals'
        principalId = $null
        resourceId = $graphSp.id
        scope = $scopeString
    } | ConvertTo-Json -Depth 5
    $newGrant = Invoke-GraphJson -Method POST -Url 'https://graph.microsoft.com/v1.0/oauth2PermissionGrants' -Body $grantBody | ConvertFrom-Json
    Write-Host "Created delegated permission grant: $($newGrant.id)"
}

[pscustomobject]@{
    displayName = $displayName
    appId = $appId
    objectId = $appObjectId
    servicePrincipalId = $clientSp.id
    scopes = $scopeString
} | ConvertTo-Json -Depth 5
