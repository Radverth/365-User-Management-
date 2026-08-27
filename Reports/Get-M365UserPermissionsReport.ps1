#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Groups

<#
.SYNOPSIS
    Lists every user in the tenant with the groups they belong to.

.DESCRIPTION
    Covers members and guests alike, and reports EFFECTIVE access - memberships
    inherited through nested groups are included.

    That makes it the right report for a permissions review, and the wrong one to
    feed into a migration. Export-GuestPermissions.ps1 deliberately records direct
    memberships only, because nesting re-inherits in the target tenant and
    replaying transitive membership would create incorrect direct memberships.

    Changes nothing. Read-only.

.PARAMETER OutputPath
    Where to write the report.

.PARAMETER TenantId
    Optional tenant ID or domain to sign in against, useful when your account
    exists in more than one tenant.

.EXAMPLE
    .\Get-M365UserPermissionsReport.ps1

.EXAMPLE
    .\Get-M365UserPermissionsReport.ps1 -OutputPath C:\Audit\permissions.csv `
        -TenantId contoso.onmicrosoft.com
#>

[CmdletBinding()]
param(
    [string]$OutputPath = ".\M365_User_Permissions_Report.csv",

    [string]$TenantId
)

$ErrorActionPreference = 'Stop'

# Ensure the parent directory exists (bare file names have no parent)
$ParentDir = Split-Path -Path $OutputPath -Parent

if ($ParentDir -and -not (Test-Path -Path $ParentDir)) {
    New-Item -ItemType Directory -Path $ParentDir -Force | Out-Null
}

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

$ConnectParams = @{
    Scopes = @('User.Read.All', 'Group.Read.All', 'Directory.Read.All')
}

if ($TenantId) { $ConnectParams['TenantId'] = $TenantId }

Connect-MgGraph @ConnectParams

# Retrieve all users
Write-Host "Retrieving users..." -ForegroundColor Cyan

# Wrap in @() so a single-user tenant still yields a countable array
$Users = @(
    Get-MgUser `
        -All `
        -Property Id,DisplayName,UserPrincipalName,Mail,AccountEnabled,UserType
)

if ($Users.Count -eq 0) {
    Write-Host "No users found in the tenant." -ForegroundColor Yellow
    Disconnect-MgGraph
    return
}

Write-Host "Found $($Users.Count) users. Fetching memberships..." -ForegroundColor Cyan

# Create report collection
$Report = [System.Collections.Generic.List[PSCustomObject]]::new()

# Cache group lookups so each group is only fetched from Graph once
$GroupCache = @{}

$Counter = 0

foreach ($User in $Users) {

    $Counter++

    Write-Progress `
        -Activity "Processing Users" `
        -Status "Processing $Counter of $($Users.Count) - $($User.DisplayName)" `
        -PercentComplete (($Counter / $Users.Count) * 100)

    try {

        # Get all direct and transitive memberships
        $Memberships = Get-MgUserTransitiveMemberOf `
            -UserId $User.Id `
            -All `
            -ErrorAction Stop

        # Keep only groups
        $GroupMemberships = @(
            $Memberships | Where-Object {
                $_.ODataType -eq "#microsoft.graph.group" -or
                $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group"
            }
        )

        if ($GroupMemberships.Count -gt 0) {

            foreach ($Group in $GroupMemberships) {

                if ($GroupCache.ContainsKey($Group.Id)) {

                    $CachedGroup = $GroupCache[$Group.Id]

                    $GroupName = $CachedGroup.GroupName
                    $GroupId   = $CachedGroup.GroupId
                    $GroupType = $CachedGroup.GroupType

                }
                else {

                    try {

                        $GroupDetails = Get-MgGroup `
                            -GroupId $Group.Id `
                            -Property Id,DisplayName,GroupTypes,SecurityEnabled `
                            -ErrorAction Stop

                        $GroupName = $GroupDetails.DisplayName
                        $GroupId   = $GroupDetails.Id

                        if ($GroupDetails.GroupTypes -contains "Unified") {
                            $GroupType = "Microsoft 365 Group / Team"
                        }
                        elseif ($GroupDetails.SecurityEnabled -eq $true) {
                            $GroupType = "Security Group"
                        }
                        else {
                            $GroupType = "Distribution Group"
                        }

                    }
                    catch {

                        $GroupName = if ($Group.AdditionalProperties["displayName"]) {
                            $Group.AdditionalProperties["displayName"]
                        }
                        else {
                            $Group.Id
                        }

                        $GroupId   = $Group.Id
                        $GroupType = "Unknown Group Type"
                    }

                    $GroupCache[$Group.Id] = [PSCustomObject]@{
                        GroupName = $GroupName
                        GroupId   = $GroupId
                        GroupType = $GroupType
                    }
                }

                $Report.Add([PSCustomObject]@{
                    DisplayName       = $User.DisplayName
                    UserPrincipalName = $User.UserPrincipalName
                    Email             = $User.Mail
                    UserType          = $User.UserType
                    AccountStatus     = if ($User.AccountEnabled) { "Enabled" } else { "Disabled" }
                    GroupName         = $GroupName
                    GroupId           = $GroupId
                    GroupType         = $GroupType
                })
            }
        }
        else {

            # User has no group memberships
            $Report.Add([PSCustomObject]@{
                DisplayName       = $User.DisplayName
                UserPrincipalName = $User.UserPrincipalName
                Email             = $User.Mail
                UserType          = $User.UserType
                AccountStatus     = if ($User.AccountEnabled) { "Enabled" } else { "Disabled" }
                GroupName         = "None"
                GroupId           = "None"
                GroupType         = "None"
            })
        }
    }
    catch {
        Write-Host "Error processing $($User.UserPrincipalName): $_" -ForegroundColor Red
    }
}

Write-Progress -Activity "Processing Users" -Completed

# Export report
Write-Host "Exporting report..." -ForegroundColor Green

$Report |
    Sort-Object DisplayName, GroupName |
    Export-Csv `
        -Path $OutputPath `
        -NoTypeInformation `
        -Encoding UTF8

Write-Host "Report exported to:" -ForegroundColor Green
Write-Host (Resolve-Path -Path $OutputPath).Path -ForegroundColor Green

# Disconnect
Disconnect-MgGraph

Write-Host "Finished successfully." -ForegroundColor Green
