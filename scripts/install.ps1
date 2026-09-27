<#
.SYNOPSIS
    Installs or updates Outlook Calendar Widget from a release zip.
.DESCRIPTION
    Run it from the folder you extracted the release zip into:

        powershell -ExecutionPolicy Bypass -File .\install.ps1

    The script:
      1. Trusts the certificate the package is signed with (OutlookCalendarWidget.cer) on this PC. The first time,
         Windows asks for administrator approval.
      2. Installs the Windows App Runtime from the Dependencies folder if this PC doesn't have it yet.
      3. Installs or updates the app for the current user. Updates keep your settings.
      4. Restarts the Widgets board so it picks up the widget.
.PARAMETER NoWidgetsRestart
    Don't restart the Widgets board after installing.
#>
[CmdletBinding()]
param(
    [switch] $NoWidgetsRestart
)

$ErrorActionPreference = 'Stop'

# The Appx cmdlets are only reliable in Windows PowerShell, so hand over to it when started from PowerShell 7.
if ($PSVersionTable.PSEdition -eq 'Core') {
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
    if ($NoWidgetsRestart) {
        $arguments += '-NoWidgetsRestart'
    }
    & "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" @arguments
    exit $LASTEXITCODE
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-PackageManifest([string] $Path) {
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $archive.GetEntry('AppxManifest.xml')
        if (-not $entry) {
            throw "$Path isn't an MSIX package."
        }
        $reader = New-Object IO.StreamReader($entry.Open())
        try {
            # The comma stops PowerShell from enumerating the document's child nodes.
            , ([xml] $reader.ReadToEnd())
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Test-Administrator {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Stop-ProcessByName([string[]] $Names) {
    foreach ($process in Get-Process -Name $Names -ErrorAction SilentlyContinue) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
}

$packages = @(Get-ChildItem -Path $PSScriptRoot -Filter 'OutlookCalendarWidget_*.msix')
if ($packages.Count -ne 1) {
    throw "Expected one OutlookCalendarWidget_*.msix next to this script but found $($packages.Count). Extract the whole release zip and run install.ps1 from there."
}
$package = $packages[0]
$manifest = Get-PackageManifest $package.FullName
$identity = $manifest.Package.Identity
$architecture = $identity.ProcessorArchitecture
Write-Host "Installing $($manifest.Package.Properties.DisplayName) $($identity.Version) ($architecture)."

$existing = Get-AppxPackage -Name $identity.Name | Where-Object { $_.Publisher -eq $identity.Publisher } | Select-Object -First 1
if ($existing -and $existing.IsDevelopmentMode) {
    throw "A development build is registered from $($existing.InstallLocation). Remove it first with: Get-AppxPackage $($identity.Name) | Remove-AppxPackage"
}

# 1. Signing certificate
$certificatePath = Join-Path $PSScriptRoot 'OutlookCalendarWidget.cer'
if (-not (Test-Path $certificatePath)) {
    throw "OutlookCalendarWidget.cer not found next to this script."
}
$certificate = Get-PfxCertificate -FilePath $certificatePath
if ($certificate.Subject -ne $identity.Publisher) {
    throw "OutlookCalendarWidget.cer ($($certificate.Subject)) doesn't match the package publisher ($($identity.Publisher))."
}
$trusted = Get-ChildItem Cert:\LocalMachine\TrustedPeople, Cert:\LocalMachine\Root | Where-Object { $_.Thumbprint -eq $certificate.Thumbprint }
if ($trusted) {
    Write-Host 'The package certificate is already trusted.'
}
else {
    Write-Host "Adding the package certificate ($($certificate.Subject), thumbprint $($certificate.Thumbprint)) to Trusted People on this PC."
    $certutil = @{
        FilePath     = 'certutil.exe'
        ArgumentList = @('-addstore', 'TrustedPeople', "`"$certificatePath`"")
        WindowStyle  = 'Hidden'
        Wait         = $true
        PassThru     = $true
    }
    if (-not (Test-Administrator)) {
        Write-Host 'Windows will ask for administrator approval.'
        $certutil.Verb = 'RunAs'
    }
    $process = Start-Process @certutil
    if ($process.ExitCode -ne 0) {
        throw "Couldn't add the certificate to Trusted People (certutil exit code $($process.ExitCode))."
    }
}

# 2. Windows App Runtime
$dependencyPaths = @()
foreach ($dependency in @($manifest.Package.Dependencies.PackageDependency)) {
    $minVersion = [version] $dependency.MinVersion
    $installed = @(Get-AppxPackage -Name $dependency.Name | Where-Object { $_.Architecture -eq $architecture -and [version] $_.Version -ge $minVersion })
    if ($installed) {
        Write-Host "$($dependency.Name) $($installed[0].Version) is already installed."
        continue
    }

    $candidates = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Dependencies') -Filter '*.msix' -ErrorAction SilentlyContinue)
    $bundled = $null
    foreach ($candidate in $candidates) {
        $candidateIdentity = (Get-PackageManifest $candidate.FullName).Package.Identity
        if ($candidateIdentity.Name -eq $dependency.Name -and $candidateIdentity.ProcessorArchitecture -eq $architecture -and [version] $candidateIdentity.Version -ge $minVersion) {
            $bundled = $candidate
            break
        }
    }
    if (-not $bundled) {
        throw "This PC needs $($dependency.Name) $minVersion or later ($architecture). Install the Windows App Runtime from https://learn.microsoft.com/windows/apps/windows-app-sdk/downloads and run the script again."
    }
    Write-Host "Installing $($dependency.Name) from the Dependencies folder."
    $dependencyPaths += $bundled.FullName
}

# 3. The app
$install = @{
    Path                     = $package.FullName
    ForceApplicationShutdown = $true
}
if ($dependencyPaths.Count -gt 0) {
    $install.DependencyPath = $dependencyPaths
}
Add-AppxPackage @install

# 4. The Widgets host caches widget providers; both processes restart on demand (Win+W).
if (-not $NoWidgetsRestart) {
    Stop-ProcessByName 'WidgetBoard', 'WidgetService'
}

Write-Host ''
Write-Host "Installed $($manifest.Package.Properties.DisplayName) $($identity.Version)."
Write-Host 'Press Win+W, select + (Add widgets) and pin Outlook Calendar. The default calendar source reads classic Outlook, so keep it running.'
