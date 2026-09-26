<#
.SYNOPSIS
    Creates the Microsoft Entra app registration the widget uses to read your calendar. Requires the Azure CLI.
.DESCRIPTION
    Registers a public client application with the delegated Microsoft Graph Calendars.Read permission and the
    Web Account Manager (WAM) broker redirect URI, then prints the application (client) ID. Paste that ID into the
    app's Settings, or bake it in with: .\scripts\deploy.ps1 -GraphClientId <id>

    Sign in to the directory that should own the registration first (az login --allow-no-subscriptions).
    Use -WhatIf to preview the request without creating anything.
.PARAMETER DisplayName
    Name shown on the consent prompt and in the Entra admin center.
.PARAMETER Audience
    Which accounts can sign in. The default allows work, school and personal Microsoft accounts.
.EXAMPLE
    .\scripts\register-entra-app.ps1
.EXAMPLE
    .\scripts\register-entra-app.ps1 -Audience AzureADMyOrg -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $DisplayName = 'Outlook Calendar Widget',

    [ValidateSet('AzureADandPersonalMicrosoftAccount', 'AzureADMultipleOrgs', 'AzureADMyOrg')]
    [string] $Audience = 'AzureADandPersonalMicrosoftAccount'
)

$ErrorActionPreference = 'Stop'

$microsoftGraph = '00000003-0000-0000-c000-000000000000'
$calendarsRead = '465a38f9-76ea-45b9-9f34-9e8b0d4b0b42' # Delegated Calendars.Read

$application = [ordered]@{
    displayName            = $DisplayName
    signInAudience         = $Audience
    isFallbackPublicClient = $true
    # Personal Microsoft accounts require v2 access tokens.
    api                    = @{ requestedAccessTokenVersion = 2 }
    requiredResourceAccess = @(
        @{
            resourceAppId  = $microsoftGraph
            resourceAccess = @(@{ id = $calendarsRead; type = 'Scope' })
        }
    )
}

function Invoke-Graph([string] $Method, [string] $Uri, [object] $Body) {
    # Pass JSON through a file: quoting JSON on the az.cmd command line is unreliable on Windows.
    $file = New-TemporaryFile
    try {
        [System.IO.File]::WriteAllText($file.FullName, ($Body | ConvertTo-Json -Depth 8))
        $output = az rest --method $Method --uri $Uri --headers 'Content-Type=application/json' --body "@$file" --only-show-errors
        if ($LASTEXITCODE -ne 0) {
            throw "az rest $Method $Uri failed (exit code $LASTEXITCODE)."
        }
        if ($output) { $output | ConvertFrom-Json }
    }
    finally {
        Remove-Item $file -ErrorAction SilentlyContinue
    }
}

if (-not $PSCmdlet.ShouldProcess("Microsoft Entra ID", "Create app registration '$DisplayName' ($Audience) with delegated Calendars.Read")) {
    $application | ConvertTo-Json -Depth 8 | Write-Host
    return
}

if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    throw 'The Azure CLI is required: https://learn.microsoft.com/cli/azure/install-azure-cli-windows'
}

az account show --only-show-errors --output none 2>$null
if ($LASTEXITCODE -ne 0) {
    throw 'Sign in first: az login --allow-no-subscriptions'
}

$created = Invoke-Graph POST 'https://graph.microsoft.com/v1.0/applications' $application
$clientId = $created.appId

# The broker redirect URI embeds the client ID, so it can only be added once the ID exists.
Invoke-Graph PATCH "https://graph.microsoft.com/v1.0/applications/$($created.id)" @{
    publicClient = @{ redirectUris = @("ms-appx-web://microsoft.aad.brokerplugin/$clientId") }
} | Out-Null

Write-Host ''
Write-Host "Created '$DisplayName'."
Write-Host "Application (client) ID: $clientId"
Write-Host ''
Write-Host 'Next: open Outlook Calendar from Start, expand Settings, paste the client ID, select Save and then Sign in.'
Write-Host "Or bake it in: .\scripts\deploy.ps1 -GraphClientId $clientId"
