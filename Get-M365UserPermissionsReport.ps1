#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Groups

# Define export path
$ExportPath = "M365_User_Permissions_Report.csv"

# Ensure the parent directory exists
$ParentDir = Split-Path -Path $ExportPath -Parent

if (-not (Test-Path -Path $ParentDir)) {
    New-Item -ItemType Directory -Path $ParentDir -Force | Out-Null
}

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

Connect-MgGraph -Scopes `
    "User.Read.All", `
    "Group.Read.All", `
    "Directory.Read.All"

# Retrieve all users
Write-Host "Retrieving users..." -ForegroundColor Cyan

$Users = Get-MgUser `
    -All `
    -Property Id,DisplayName,UserPrincipalName,Mail,AccountEnabled,UserType

if (-not $Users) {
    Write-Host "No users found in the tenant." -ForegroundColor Yellow
    Disconnect-MgGraph
    return
}

Write-Host "Found $($Users.Count) users. Fetching memberships..." -ForegroundColor Cyan

# Create report collection
$Report = [System.Collections.Generic.List[PSCustomObject]]::new()

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
        $GroupMemberships = $Memberships | Where-Object {
            $_.ODataType -eq "#microsoft.graph.group" -or
            $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group"
        }

        if ($GroupMemberships) {

            foreach ($Group in $GroupMemberships) {

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

# Export report
Write-Host "Exporting report..." -ForegroundColor Green

$Report |
    Sort-Object DisplayName, GroupName |
    Export-Csv `
        -Path $ExportPath `
        -NoTypeInformation `
        -Encoding UTF8

Write-Host "Report exported to:" -ForegroundColor Green
Write-Host $ExportPath -ForegroundColor Green

# Disconnect
Disconnect-MgGraph

Write-Host "Finished successfully." -ForegroundColor Green
