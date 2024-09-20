# Add the SharePoint PowerShell snap-in to enable SharePoint cmdlets
Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

#region === Configurable Variables ===

# Replace these URLs with the actual site and web URLs you want to process
  $SiteURLs = @("https://rootsite1", "https://rootsite2")
  $WebURLs = @("https://rootsite1/subweb1", "https://rootsite2/subweb2")
  $IncludeItems = $false  # Set to $true to include item-level permissions

# Email configuration
  $msgFrom = "NoReply@domain.com"
  $msgTo = "first.last@domain.com"  # Replace with appropriate address
  $SmtpServer = "mxrelay.domain.com"
  $msgSubject = "SharePoint Permissions Report"
  $msgAttachment = "C:\SitePermissionReport.html"  # Update the path if necessary

#endregion === Configurable Variables ===

# Define sensitive groups to highlight in the report
$SensitiveGroups = @("Everyone", "NT AUTHORITY\Authenticated Users", "All Users")

# HTML report components with improved styling
$reportHeader = @'
<!DOCTYPE html>
<html>
<head>
<style>
body {
    font-family: Verdana, Arial, sans-serif;
    background-color: #f2f2f2;
}
table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 20px;
}
th, td {
    border: 1px solid #999;
    padding: 8px;
    text-align: left;
}
th {
    background-color: #1f4e79;
    color: white;
}
tr:nth-child(even) {
    background-color: #f9f9f9;
}
.warning {
    color: red;
    font-weight: bold;
}
</style>
</head>
<body>
'@

$reportBody = ""
$reportFooter = "</body></html>"

# === Functions ===

# Function to get permissions for a site collection (SPSite)
function Get-PermissionsForSPSite {
    Param(
        [Microsoft.SharePoint.SPSite]$SPSite,
        [bool]$IncludeItems = $false
    )

    Write-Host "Processing Site Collection: $($SPSite.Url)"

    $report = @"
<h1>Permissions Report for Site Collection: $($SPSite.Url)</h1>
"@

    # Iterate through all webs in the site collection
    foreach ($SPWeb in $SPSite.AllWebs) {
        try {
            $report += Get-PermissionsForSPWeb -SPWeb $SPWeb -IncludeItems:$IncludeItems
        }
        finally {
            $SPWeb.Dispose()
        }
    }

    return $report
}

# Function to get permissions for a specific web (SPWeb)
function Get-PermissionsForSPWeb {
    Param(
        [Microsoft.SharePoint.SPWeb]$SPWeb,
        [bool]$IncludeItems = $false
    )

    Write-Host "Processing Web: $($SPWeb.Url)"

    $report = @"
<h1>Permissions Report for Web: $($SPWeb.Url)</h1>
"@

    # Process the specified web
    $report += ProcessSPWeb -SPWeb $SPWeb -IncludeItems:$IncludeItems

    # Add group memberships assigned to this web
    $report += Get-GroupMemberships -SPWeb $SPWeb

    return $report
}

# Helper function to process an SPWeb object
function ProcessSPWeb {
    Param(
        [Microsoft.SharePoint.SPWeb]$SPWeb,
        [bool]$IncludeItems = $false
    )

    Write-Host "Processing Web: $($SPWeb.Title) ($($SPWeb.Url))"

    $report = @"
<h2>Permissions for Web: $($SPWeb.Title) ($($SPWeb.Url))</h2>
"@

    # Get permissions for the current web
    if ($SPWeb.HasUniqueRoleAssignments) {
        $report += Get-ObjectPermissions -SPObject $SPWeb -ObjectType "Web"
    } else {
        $report += "<p>This web inherits permissions from its parent.</p>"
    }

    # Iterate through all lists and libraries
    foreach ($list in $SPWeb.Lists) {
        Write-Host "  Checking List/Library: $($list.Title)"

        # Start the section for the list/library
        $report += @"
<h3>List/Library: $($list.Title)</h3>
"@

        if ($list.HasUniqueRoleAssignments) {
            $report += Get-ObjectPermissions -SPObject $list -ObjectType "List"
        } else {
            $report += "<p>This list/library inherits permissions from its parent.</p>"
        }

        # Optionally include item-level permissions
        if ($IncludeItems) {
            # Only process items with unique permissions
            $uniqueItems = $list.Items | Where-Object { $_.HasUniqueRoleAssignments -eq $true }
            foreach ($item in $uniqueItems) {
                Write-Host "    Checking Item: $($item.Name)"

                # Start the section for the item
                $report += @"
<h4>Item: $($item.Name)</h4>
"@

                $report += Get-ObjectPermissions -SPObject $item -ObjectType "Item"
            }
        }
    }

    return $report
}

# Function to get permissions for a securable object
function Get-ObjectPermissions {
    Param(
        [Microsoft.SharePoint.ISecurableObject]$SPObject,
        [string]$ObjectType
    )

    $roleAssignments = $SPObject.RoleAssignments
    $objectUrl = $SPObject.Url

    # Get the title if available (for lists and items)
    $objectTitle = ""
    if ($ObjectType -eq "List" -or $ObjectType -eq "Item") {
        $objectTitle = $SPObject.Title
    }

    $reportSection = @"
<table>
    <tr>
        <th colspan="2">$ObjectType : $objectTitle ($objectUrl)</th>
    </tr>
    <tr>
        <th>User/Group</th>
        <th>Permissions</th>
    </tr>
"@

    foreach ($roleAssignment in $roleAssignments) {
        $member = $roleAssignment.Member
        $roles = $roleAssignment.RoleDefinitionBindings
        $roleNames = ($roles | Select-Object -ExpandProperty Name) -join ", "

        # Check for sensitive groups
        if ($SensitiveGroups -contains $member.Name) {
            $memberName = "<span class='warning'>$($member.Name)</span>"
        } else {
            $memberName = $member.Name
        }

        $reportSection += @"
    <tr>
        <td>$memberName</td>
        <td>$roleNames</td>
    </tr>
"@
    }

    $reportSection += "</table><br>"
    return $reportSection
}

# Function to get groups assigned to the web and their members
function Get-GroupMemberships {
    Param(
        [Microsoft.SharePoint.SPWeb]$SPWeb
    )

    Write-Host "Collecting group memberships for web: $($SPWeb.Url)"

    $reportSection = "<h2>Groups Assigned to Web: $($SPWeb.Title) ($($SPWeb.Url))</h2>"

    # Get groups that have permissions on the web
    $groups = @()
    foreach ($roleAssignment in $SPWeb.RoleAssignments) {
        $member = $roleAssignment.Member
        if ($member -is [Microsoft.SharePoint.SPGroup]) {
            $groups += $member
        }
    }

    if ($groups.Count -gt 0) {
        $reportSection += "<table>
            <tr>
                <th>Group Name</th>
                <th>Members</th>
            </tr>"

        foreach ($group in $groups) {
            $reportSection += "<tr>
                <td>$($group.Name)</td>
                <td><ul>"

            foreach ($user in $group.Users) {
                $reportSection += "<li>$($user.Name) ($($user.LoginName))</li>"
            }

            $reportSection += "</ul></td></tr>"
        }

        $reportSection += "</table><br>"
    } else {
        $reportSection += "<p>No groups have permissions assigned to this web.</p>"
    }

    return $reportSection
}

# Function to send the report via email
function Send-Report {
    # Write the report to the file
    $fullReport = $reportHeader + $reportBody + $reportFooter
    $fullReport | Out-File -FilePath $msgAttachment -Encoding UTF8

    # Email body
    $msgBody = @'
This email was sent from an unattended account and is for information purposes only.<br><br>
Please find attached the SharePoint permissions report.<br><br>
Thank you,<br>
SharePoint Administrators
'@

    # Send the email
    Send-MailMessage -To $msgTo -From $msgFrom -Subject $msgSubject -Body $msgBody -BodyAsHtml -Attachments $msgAttachment -SmtpServer $SmtpServer

    # Clean up
    Remove-Item $msgAttachment
}

# Main function to coordinate the execution
function Main {
    # Record the start time
    $scriptStartTime = Get-Date
    Write-Host "Script started at: $($scriptStartTime)"

    # Initialize report body
    $reportBody = ""

    # Process Site URLs
    foreach ($siteUrl in $SiteURLs) {
        $SPSite = Get-SPSite $siteUrl
        if ($SPSite -ne $null) {
            try {
                $reportBody += Get-PermissionsForSPSite -SPSite $SPSite -IncludeItems:$IncludeItems
            }
            catch {
                Write-Error "Error processing site $siteUrl : $_"
            }
            finally {
                $SPSite.Dispose()
            }
        }
        else {
            Write-Warning "Could not access site $siteUrl"
        }
    }

    # Process Web URLs
    foreach ($webUrl in $WebURLs) {
        $SPWeb = Get-SPWeb $webUrl
        if ($SPWeb -ne $null) {
            try {
                $reportBody += Get-PermissionsForSPWeb -SPWeb $SPWeb -IncludeItems:$IncludeItems
            }
            catch {
                Write-Error "Error processing web $webUrl : $_"
            }
            finally {
                $SPWeb.Dispose()
            }
        }
        else {
            Write-Warning "Could not access web $webUrl"
        }
    }

    # Send the report via email
    if ($reportBody -ne "") {
        Send-Report
        Write-Host "Report sent successfully."
    }
    else {
        Write-Host "No data to report."
    }

    # Record the end time and calculate elapsed time
    $scriptEndTime = Get-Date
    Write-Host "Script ended at: $($scriptEndTime)"
    $elapsedTime = $scriptEndTime - $scriptStartTime
    Write-Host "Total elapsed time: $($elapsedTime.Hours) hours, $($elapsedTime.Minutes) minutes, $($elapsedTime.Seconds) seconds."
}

# Run the main function
Main
