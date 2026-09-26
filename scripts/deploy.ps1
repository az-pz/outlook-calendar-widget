<#
.SYNOPSIS
    Builds the app, registers it for the current user and restarts the Widgets host so it picks up the change.
.DESCRIPTION
    Registers the build output as a loose-file MSIX package (requires Developer Mode), so there is no certificate to create
    or trust. Re-running the script after code changes updates the installed app in place.
.PARAMETER Platform
    x64 or ARM64. Defaults to the architecture of this machine.
.PARAMETER Configuration
    Debug (default) or Release.
.PARAMETER GraphClientId
    Optional Microsoft Entra application (client) ID to bake in as the default, so users don't have to paste it in Settings.
.PARAMETER NoWidgetsRestart
    Skip restarting the Widgets board/service after registering. New or changed widget definitions won't show up until
    they restart.
.EXAMPLE
    .\scripts\deploy.ps1
.EXAMPLE
    .\scripts\deploy.ps1 -Configuration Release -GraphClientId 00000000-0000-0000-0000-000000000000
#>
[CmdletBinding()]
param(
    [ValidateSet('x64', 'ARM64')]
    [string] $Platform = $(if ([System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture -eq 'Arm64') { 'ARM64' } else { 'x64' }),

    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Debug',

    [string] $GraphClientId,

    [switch] $NoWidgetsRestart
)

$ErrorActionPreference = 'Stop'
$project = Join-Path $PSScriptRoot '..\src\OutlookCalendarWidget\OutlookCalendarWidget.csproj' | Resolve-Path
$projectDir = Split-Path $project

function Stop-ProcessByName([string[]] $Names) {
    foreach ($process in Get-Process -Name $Names -ErrorAction SilentlyContinue) {
        Write-Host "Stopping $($process.ProcessName) ($($process.Id))"
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
}

$devMode = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock' -Name AllowDevelopmentWithoutDevLicense -ErrorAction SilentlyContinue
if ($devMode.AllowDevelopmentWithoutDevLicense -ne 1) {
    throw 'Developer Mode is off. Turn it on in Settings > System > For developers, then run this script again.'
}

# The running widget provider locks the build output, and the Widgets host memory-maps the package's resources.pri.
# All of them restart on demand (Win+W).
Stop-ProcessByName 'OutlookCalendarWidget', 'WidgetBoard', 'WidgetService'

$buildArgs = @('build', $project, '-c', $Configuration, "-p:Platform=$Platform", '-nologo', '-v:minimal')
if ($GraphClientId) {
    $buildArgs += "-p:GraphClientId=$GraphClientId"
}
& dotnet @buildArgs
if ($LASTEXITCODE -ne 0) {
    throw "Build failed (exit code $LASTEXITCODE)."
}

$manifest = Get-ChildItem (Join-Path $projectDir "bin\$Platform\$Configuration") -Filter AppxManifest.xml -Recurse |
    Sort-Object { $_.FullName.Length } |
    Select-Object -First 1
if (-not $manifest) {
    throw "AppxManifest.xml not found under bin\$Platform\$Configuration."
}

# Loose-file registration can't install frameworks, so check for the Windows App Runtime up front.
[xml] $xml = Get-Content $manifest.FullName
foreach ($dependency in $xml.Package.Dependencies.PackageDependency) {
    $installed = Get-AppxPackage -Name $dependency.Name |
        Where-Object { [version] $_.Version -ge [version] $dependency.MinVersion -and $_.Architecture -eq $Platform }
    if (-not $installed) {
        throw "Missing dependency $($dependency.Name) >= $($dependency.MinVersion) ($Platform). Install the matching Windows App Runtime from https://learn.microsoft.com/windows/apps/windows-app-sdk/downloads."
    }
}

Write-Host "Registering $($manifest.FullName)"
Add-AppxPackage -Register $manifest.FullName -ForceApplicationShutdown

if (-not $NoWidgetsRestart) {
    # The Widgets host caches provider definitions; both processes restart on demand (Win+W).
    Stop-ProcessByName 'WidgetBoard', 'WidgetService'
}

Write-Host ''
Write-Host 'Done. Keep classic Outlook running (or pick another calendar source in the app), then press Win+W and choose + to add the widget.'
