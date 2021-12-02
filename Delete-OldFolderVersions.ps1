<#
    .SYNOPSIS
     Identifies, reports and deletes old OWA / ECP folders
    .DESCRIPTION
     Run this Script to list old and unused OWA / ECP folder. Set Parameter -DeleteOldVersions to $true if you
     want to delete old folders to free up disk space.
    .PARAMETER DeleteOldVersions
     Set DeleteOldVersions to $True to delete old OWA / ECP folder versions.
    .EXAMPLE
     .\Delete-OldFolderVersions.ps1
     List old OWA / ECP folderversions:
    .EXAMPLE
     .\Delete-OldFolderVersions.ps1 -DeleteOldVersions $true
     List and delete old OWA / ECP Directory versions:
    .NOTES
     Author:  Frank Zoechling
     Website: https://www.frankysweb.de
     Twitter: @FrankysWeb
#>
 
Param(
    [Parameter(Mandatory=$False)]
    [bool]$DeleteOldVersions
)
 
# Test if evelated Shell
Function Confirm-Administrator {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
    if ($currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
        return $true
    } else {
        return $false
    }
}
 
if (-not (Confirm-Administrator)) {
    Write-Output $msgNewLine
    Write-Warning "This script needs to be executed in elevated mode. Start the Exchange Management Shell as an Administrator and try again."
    $Error.Clear()
    Start-Sleep -Seconds 2
    exit
}
 
#Foldernames to search for old ECP / OWA folders
$ECPFolderPath = $exinstall + "ClientAccess\ecp"
$OWAFolderPath = $exinstall + "ClientAccess\Owa"
$OWAPremFolderPath = $exinstall + "ClientAccess\Owa\prem"
 
#Get Exchange Server Versions
try {
    $ExchangeServerVersions = @()
    $ExchangeServerDisplayVersions = (Get-ExchangeServer | Where-Object { $_.AdminDisplayVersion.Major -eq 15 }).AdminDisplayVersion
    foreach ($ExchangeServerDisplayVersion in $ExchangeServerDisplayVersions) {
        $ExchangeServerVersions += $ExchangeServerDisplayVersion.Major,$ExchangeServerDisplayVersion.Minor,$ExchangeServerDisplayVersion.Build,$ExchangeServerDisplayVersion.Revision -join "."
    }
    $ExchangeServerVersion = $ExchangeServerVersions | Sort-Object | Select-Object -first 1
    write-host ""
    write-host "Oldest installed Exchange Build is: $ExchangeServerVersion"
    write-host ""
} catch {}
 
 
#Search all OWA / ECP folder versions
try {
    $AllVersions = @()
    $AllVersions += Get-ChildItem $ECPFolderPath -Directory | Where-Object { $_.Name.StartsWith("15.") } | Select-Object FullName,Name
    $AllVersions += Get-ChildItem $OWAFolderPath -Directory | Where-Object { $_.Name.StartsWith("15.") } | Select-Object FullName,Name
    $AllVersions += Get-ChildItem $OWAPremFolderPath -Directory | Where-Object { $_.Name.StartsWith("15.") } | Select-Object FullName,Name
} catch {}
 
#Search for old versions
write-host "Old OWA / ECP folders:"
write-host ""
[int]$ExchangeBuild = $ExchangeServerVersion.Replace(".","")
$OldVersions = @()
foreach ($Version in $AllVersions)  {
    [int]$Folderversion = $Version.Name.Replace(".","")
    if ( $Folderversion -lt $ExchangeBuild ) {
        write-host $Version.Fullname
        $OldVersions += $Version.Fullname
    }
}
write-host ""
 
#Delete old OWA / ECP folder versions
if ($DeleteOldVersions -eq $true) {
    write-host "Deleting old OWA / ECP versions:"
    write-host ""
    foreach ($Oldversion in $OldVersions) {
        write-host $Oldversion
        Remove-Item $Oldversion -Recurse -Confirm:$true
    }
}