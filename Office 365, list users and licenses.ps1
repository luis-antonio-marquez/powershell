<#
.SYNOPSIS
Export list of users and assigned licenses in a tenant, using the module MSOnline
It is an interactive program, not a module.

The connection to the MSOL service must be stablished outside of the function.

.DESCRIPTION

.PARAMETER

.EXAMPLE
    Import-Module MSOnline
    $UserCredential = Get-Credential
    Connect-MsolService -Credential $UserCredential

    Clear-Host

    [string]$p_csvFile = "c:\tmp\Office365_UserExport_" + (Get-Date).ToString('yyyyMMdd_HHmmss') + ".csv"
    [string]$p_delimiter = ","
    [string]$p_tennantNameToBeRemovedFromTheLicenseStrings = "myTenantName"

    Export-O365UsersAndLicenses `
        -csvFile $p_csvFile `
        -delimiter $p_delimiter `
        -show_progress $true `
        -remove_string $p_tennantNameToBeRemovedFromTheLicenseStrings

    Write-Host "> Output File: $p_csvFile " -ForegroundColor Yellow
    Write-Host '> END' -ForegroundColor Yellow
.NOTES
Initial - June 2020
#>

function Export-O365UsersAndLicenses {
    Param(
        [Parameter (Mandatory = $true)] [string]$csvFile = $null,
        [Parameter (Mandatory = $true)] [string]$delimiter = $null,
        [Parameter (Mandatory = $true)] [bool]$show_progress = $true,
        [Parameter (Mandatory = $true)] [string]$remove_string = $null
    )

    # Assign default values
    if ($csvFile -eq $null) {
        $csvFile = "UserExport_" + (Get-Date).ToString('yyyyMMdd_HHmmss') + ".csv"
    }
    if ($delimiter -eq $null) {
        $delimiter = ','
    }

    if (Test-Path $csvFile) { Remove-Item $csvFile }

    $users = get-MSOLUser -All -EnabledFilter EnabledOnly

    foreach ( $u in $users) {

        $licenses = ''
        foreach ($l in $u.licenses) {
            $licenses = $licenses + ($l.AccountSkuId.ToString())
        }
        if ($remove_string -ne $null) {
            $licenses = $licenses.Replace($remove_string, '')
        }

        $u  | Select-Object  UserPrincipalName, DisplayName, UserType, Department, Office, UsageLocation, Title, IsLicensed, @{N = 'Licenses'; E = { $licenses } },  LastDirSyncTime, LastPasswordChangeTimestamp, ValidationStatus, WhenCreated, PasswordNeverExpires `
            | Export-Csv -Path $csvFile -Delimiter $delimiter -Encoding 'UTF8' -Append

        if ($show_progress) {
            Write-Host "Exporting $($u.UserPrincipalName) ..." -ForegroundColor Green
        }
    }

}


# ------------ Main / Example
<# Import-Module MSOnline
$UserCredential = Get-Credential
Connect-MsolService -Credential $UserCredential

Clear-Host

[string]$p_csvFile = "c:\tmp\Office365_UserExport_" + (Get-Date).ToString('yyyyMMdd_HHmmss') + ".csv"
[string]$p_delimiter = ","
[string]$p_tennantNameToBeRemovedFromTheLicenseStrings = "myTenantName"

Export-O365UsersAndLicenses `
    -csvFile $p_csvFile `
    -delimiter $p_delimiter `
    -show_progress $true `
    -remove_string $p_tennantNameToBeRemovedFromTheLicenseStrings

Write-Host "> Output File: $p_csvFile " -ForegroundColor Yellow
Write-Host '> END' -ForegroundColor Yellow #>


