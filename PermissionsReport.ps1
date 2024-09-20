# Add the SharePoint PowerShell snap-in to enable SharePoint cmdlets
Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

#region === Configurable Variables ===

# Replace these URLs with the actual site and web URLs you want to process
  $SiteURLs = @("https://rootsite1", "https://rootsite2")
  $WebURLs = @("https://rootsite1/subweb1", "https://rootsite2/subweb2")
  $IncludeItems = $true

# Email configuration
  $msgFrom = "NoReply@domain.com"
  $msgTo = "first.last@domain.com"  # Replace with appropriate address
  $SmtpServer = "mxrelay.domain.com"
  $msgSubject = "SharePoint Permissions Report"
  $msgAttachment = "C:\SitePermissionReport.html"  # Update the path if necessary

#endregion === Configurable Variables ===

# Define sensitive groups to highlight in the report
$SensitiveGroups = @(
    "everyone",
    "nt authority\\authenticated users",
    "all users",
    "all authenticated users",
    "authenticated users"
)


# Initialize global hashtable to track listed groups
$global:ListedGroups = @{}


# Mapping of special claims identifiers to friendly names
$global:SpecialGroups = @{
    "c:0(.s|true" = "All Authenticated Users";
    "c:0!.s|windows" = "All Users (Windows)";
    # Add more mappings as needed
}


# Initialize the log file
$LogFilePath = "C:\PermissionsReportLog.txt"
if (Test-Path $LogFilePath) {
    Remove-Item $LogFilePath
}


# Function to log messages with anonymization
function Log-Message {
    param (
        [string]$Message,
        [string]$Level = "Info"
    )


    # Anonymize sensitive data
    $anonymizedMessage = $Message -replace 'https?://\S+', '[URL]' `
                                    -replace '[Uu]ser:\s*\S+', 'User:[REDACTED]' `
                                    -replace '[Ll]ist:\s*\S+', 'List:[REDACTED]' `
                                    -replace '[Ww]eb:\s*\S+', 'Web:[REDACTED]'


    # Write to log file
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Level] $anonymizedMessage" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
}


# HTML report components with improved styling and JavaScript
$reportHeader = @'
<!DOCTYPE html>
<html>
<head>
<style>
body {
    font-family: Verdana, Arial, sans-serif;
    background-color: #f2f2f2;
    margin: 0;
    padding: 20px;
}
h1 {
    color: #1f4e79;
    border-bottom: 2px solid #1f4e79;
    padding-bottom: 10px;
}
h2 {
    color: #1f4e79;
    margin-top: 30px;
}
h3 {
    color: #1f4e79;
    margin-top: 20px;
}
h4 {
    color: #1f4e79;
    margin-top: 15px;
}
h5 {
    color: #1f4e79;
    margin-top: 10px;
}
.collapsible {
    margin-bottom: 15px;
    margin-left: 20px;
}
.collapsible-header {
    font-weight: bold;
    cursor: pointer;
    position: relative;
    padding-left: 20px;
}
.collapsible-header::before {
    content: "\25B6"; /* Right-pointing triangle */
    position: absolute;
    left: 0;
    transition: transform 0.2s ease-in-out;
}
.collapsible-header.active::before {
    content: "\25BC"; /* Down-pointing triangle */
}
.collapsible-content {
    display: none;
    margin-left: 20px;
}
table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 20px;
    margin-left: 20px;
}
th, td {
    border: 1px solid #999;
    padding: 8px;
    text-align: left;
    vertical-align: top;
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
<script>
document.addEventListener("DOMContentLoaded", function() {
    var headers = document.querySelectorAll(".collapsible-header");
    headers.forEach(function(header) {
        header.addEventListener("click", function() {
            this.classList.toggle("active");
            var content = this.nextElementSibling;
            if (content.style.display === "block") {
                content.style.display = "none";
            } else {
                content.style.display = "block";
            }
        });
    });
});
</script>
'@


$reportBody = ""
$reportFooter = "</body></html>"


# === Functions ===


# Function to get permissions for a site collection (SPSite)
function Get-PermissionsForSPSite {
    Param(
        [Microsoft.SharePoint.SPSite]$SPSite,
        [bool]$IncludeItems
    )


    Write-Host "Processing Site Collection: $($SPSite.Url)" -ForegroundColor Cyan
    Log-Message "Processing Site Collection: $($SPSite.Url)" "Info"


    $report = @"
<h1>Permissions Report for Site Collection: $($SPSite.Url)</h1>
"@


    # Process the root web separately
    $rootWeb = $SPSite.RootWeb
    try {
        $report += ProcessSPWeb -SPWeb $rootWeb -IncludeItems:$IncludeItems -IsRootWeb:$true
    }
    finally {
        $rootWeb.Dispose()
    }


    # Iterate through all subsites
    foreach ($SPWeb in $SPSite.AllWebs) {
        if ($SPWeb.Url -ne $SPSite.RootWeb.Url) {
            try {
                $report += Get-PermissionsForSPWeb -SPWeb $SPWeb -IncludeItems:$IncludeItems
            }
            finally {
                $SPWeb.Dispose()
            }
        }
    }


    return $report
}


# Function to get permissions for a specific web (SPWeb)
function Get-PermissionsForSPWeb {
    Param(
        [Microsoft.SharePoint.SPWeb]$SPWeb,
        [bool]$IncludeItems
    )


    Write-Host "Processing Web: $($SPWeb.Url)" -ForegroundColor Cyan
    Log-Message "Processing Web: $($SPWeb.Url)" "Info"


    $report = ""


    # Process the specified web
    $report += ProcessSPWeb -SPWeb $SPWeb -IncludeItems:$IncludeItems -IsRootWeb:$false


    return $report
}


# Helper function to process an SPWeb object
function ProcessSPWeb {
    Param(
        [Microsoft.SharePoint.SPWeb]$SPWeb,
        [bool]$IncludeItems,
        [bool]$IsRootWeb
    )


    if ($IsRootWeb) {
        $report = @"
<h2>Permissions for Site: $($SPWeb.Title) ($($SPWeb.Url))</h2>
"@
    } else {
        $report = @"
<h2>Permissions for Subsite: $($SPWeb.Title) ($($SPWeb.Url))</h2>
"@
    }


    # Reset the global ListedGroups for each web to avoid cross-web contamination
    $global:ListedGroups = @{}


    # Get permissions for the current web
    if ($SPWeb.HasUniqueRoleAssignments) {
        Write-Host "  Web '$($SPWeb.Title)' has unique permissions." -ForegroundColor Green
        Log-Message "Web '$($SPWeb.Title)' has unique permissions." "Info"
        $report += Get-ObjectPermissions -SPObject $SPWeb -ObjectType "Web"
    } else {
        Write-Host "  Web '$($SPWeb.Title)' inherits permissions from its parent." -ForegroundColor Yellow
        Log-Message "Web '$($SPWeb.Title)' inherits permissions from its parent." "Info"
        $report += "<p>This web inherits permissions from its parent.</p>"
    }


    # Categorize lists/libraries
    $listsInheriting_NoUniqueItems = @()
    $listsInheriting_WithUniqueItems = @()
    $listsWithUniquePermissions = @()


    foreach ($list in $SPWeb.Lists) {
        if ($list.Hidden -or $list.BaseType -eq "None") {
            continue  # Skip hidden or system lists
        }


        $uniqueItems = @()
        $hasUniqueItems = $false


        if ($IncludeItems) {
            try {
                $items = $list.GetItems()
                foreach ($item in $items) {
                    if ($item.HasUniqueRoleAssignments) {
                        $uniqueItems += $item
                        $hasUniqueItems = $true
                    }
                    else {
                                            $uniqueItems += $item
                        $hasUniqueItems = $true
                    }
                }
            } catch {
                Write-Host "Error accessing items in list '$($list.Title)': $_" -ForegroundColor Red
                Log-Message "Error accessing items in list '$($list.Title)': $_" "Error"
            }
        }


        if ($list.HasUniqueRoleAssignments) {
            Write-Host "  List/Library '$($list.Title)' has unique permissions." -ForegroundColor Green
            Log-Message "List/Library '$($list.Title)' has unique permissions." "Info"
            $listsWithUniquePermissions += [PSCustomObject]@{
                List = $list
                UniqueItems = $uniqueItems
            }
        } elseif ($hasUniqueItems) {
            Write-Host "  List/Library '$($list.Title)' inherits permissions but has items with unique permissions." -ForegroundColor Yellow
            Log-Message "List/Library '$($list.Title)' inherits permissions but has items with unique permissions." "Info"
            $listsInheriting_WithUniqueItems += [PSCustomObject]@{
                List = $list
                UniqueItems = $uniqueItems
            }
        } else {
            Write-Host "  List/Library '$($list.Title)' inherits permissions with no unique items." -ForegroundColor Gray
            Log-Message "List/Library '$($list.Title)' inherits permissions with no unique items." "Info"
            $listsInheriting_NoUniqueItems += $list
        }
    }


    # Display lists/libraries inheriting permissions with no unique items
    if ($listsInheriting_NoUniqueItems.Count -gt 0) {
        $report += @"
<h3>Lists/Libraries Inheriting Permissions with No Unique Items:</h3>
<table>
    <tr>
        <th>Title</th>
        <th>URL</th>
    </tr>
"@
        foreach ($list in $listsInheriting_NoUniqueItems) {
            $report += @"
    <tr>
        <td>$($list.Title)</td>
        <td><a href='$($list.DefaultViewUrl)'>$($list.DefaultViewUrl)</a></td>
    </tr>
"@
        }
        $report += "</table>"
    }


    # Display lists/libraries inheriting permissions with unique items
    if ($listsInheriting_WithUniqueItems.Count -gt 0) {
        Write-Host "Found $($listsInheriting_WithUniqueItems.Count) lists inheriting permissions with unique items." -ForegroundColor Cyan
        Log-Message "Found $($listsInheriting_WithUniqueItems.Count) lists inheriting permissions with unique items." "Info"
        $report += @"
<h3>Lists/Libraries Inheriting Permissions with Unique Items:</h3>
"@
        foreach ($item in $listsInheriting_WithUniqueItems) {
            $list = $item.List
            $uniqueItems = $item.UniqueItems


            $report += @"
<h4>$($list.Title)</h4>
"@
            $report += @"
<div class='collapsible'>
    <span class='collapsible-header'>Items with Unique Permissions in '$($list.Title)'</span>
    <div class='collapsible-content'>
"@
            foreach ($uniqueItem in $uniqueItems) {
                Write-Host "    Processing Item with unique permissions: $($uniqueItem.Name)" -ForegroundColor Green
                Log-Message "Processing Item with unique permissions: $($uniqueItem.Name)" "Info"


                $itemUrl = $uniqueItem.Url


                $report += @"
<h5>Item: $($uniqueItem.Name) ($itemUrl)</h5>
"@
                $report += Get-ObjectPermissions -SPObject $uniqueItem -ObjectType "Item"
            }
            $report += "</div></div>"
        }
    }


    # Display lists/libraries with unique permissions
    if ($listsWithUniquePermissions.Count -gt 0) {
        $report += @"
<h3>Lists/Libraries with Unique Permissions:</h3>
"@
        foreach ($item in $listsWithUniquePermissions) {
            $list = $item.List
            $uniqueItems = $item.UniqueItems


            Write-Host "  Processing List/Library with unique permissions: $($list.Title)" -ForegroundColor Green
            Log-Message "Processing List/Library with unique permissions: $($list.Title)" "Info"


            $report += @"
<h4>$($list.Title)</h4>
"@
            $report += Get-ObjectPermissions -SPObject $list -ObjectType "List"


            if ($IncludeItems -and $uniqueItems.Count -gt 0) {
                Write-Host "    Found $($uniqueItems.Count) items with unique permissions in '$($list.Title)'." -ForegroundColor Cyan
                Log-Message "Found $($uniqueItems.Count) items with unique permissions in '$($list.Title)'." "Info"
                $report += @"
<div class='collapsible'>
    <span class='collapsible-header'>Items with Unique Permissions in '$($list.Title)'</span>
    <div class='collapsible-content'>
"@
                foreach ($uniqueItem in $uniqueItems) {
                    Write-Host "    Processing Item with unique permissions: $($uniqueItem.Name)" -ForegroundColor Green
                    Log-Message "Processing Item with unique permissions: $($uniqueItem.Name)" "Info"


                    $itemUrl = $uniqueItem.Url


                    $report += @"
<h5>Item: $($uniqueItem.Name) ($itemUrl)</h5>
"@
                    $report += Get-ObjectPermissions -SPObject $uniqueItem -ObjectType "Item"
                }
                $report += "</div></div>"
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


    # Begin table without redundant title
    $reportSection = @"
<table>
    <tr>
        <th>User/Group</th>
        <th>Permissions</th>
        <th>Members</th>
    </tr>
"@


    foreach ($roleAssignment in $roleAssignments) {
        $member = $roleAssignment.Member
        $roles = $roleAssignment.RoleDefinitionBindings
        $roleNames = ($roles | Select-Object -ExpandProperty Name) -join ", "


        # Normalize the member name
        $normalizedMemberName = $member.Name.Trim().ToLower()


        # Check for sensitive groups (case-insensitive)
        if ($SensitiveGroups -contains $normalizedMemberName) {
            $memberName = "<span class='warning'>$($member.Name)</span>"
        } else {
            $memberName = $member.Name
        }


        # Initialize members column
        $membersColumn = ""


        # Handle special claims identifiers
        $loginName = $member.LoginName
        if ($global:SpecialGroups.ContainsKey($loginName)) {
            $memberName = $global:SpecialGroups[$loginName]
            $membersColumn = ""
        } elseif ($member -is [Microsoft.SharePoint.SPGroup]) {
            # If the group has not been listed yet
            if (-not $global:ListedGroups.ContainsKey($member.ID)) {
                $global:ListedGroups[$member.ID] = $true  # Mark group as listed
                $groupMembers = @()
                foreach ($user in $member.Users) {
                    try {
                        # Ensure the user to get the latest information
                        $resolvedUser = $SPObject.Site.RootWeb.EnsureUser($user.LoginName)
                        $email = $resolvedUser.Email
                        if ([string]::IsNullOrEmpty($email)) {
                            $email = $resolvedUser.LoginName
                        }
                        $groupMembers += "$($resolvedUser.Name) &lt;$email&gt;"
                    } catch {
                        $groupMembers += "$($user.Name) &lt;$($user.LoginName)&gt;"
                    }
                }
                if ($groupMembers.Count -gt 0) {
                    $membersColumn = "<ul><li>" + ($groupMembers -join "</li><li>") + "</li></ul>"
                }
            } else {
                $membersColumn = "(Already listed above)"
            }
        } elseif ($member -is [Microsoft.SharePoint.SPUser]) {
            try {
                # Ensure the user to get the latest information
                $resolvedUser = $SPObject.Site.RootWeb.EnsureUser($member.LoginName)
                $email = $resolvedUser.Email
                if ([string]::IsNullOrEmpty($email)) {
                    $email = $resolvedUser.LoginName
                }
                $membersColumn = "&lt;$email&gt;"
            } catch {
                $membersColumn = "&lt;$($member.LoginName)&gt;"
            }
        }


        $reportSection += @"
    <tr>
        <td>$memberName</td>
        <td>$roleNames</td>
        <td>$membersColumn</td>
    </tr>
"@
    }


    $reportSection += "</table><br>"
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
    #Send-MailMessage -To $msgTo -From $msgFrom -Subject $msgSubject -Body $msgBody -BodyAsHtml -Attachments $msgAttachment -SmtpServer $SmtpServer


    # Clean up
    #Remove-Item $msgAttachment
}


# Main function to coordinate the execution
function Main {
    # Record the start time
    $scriptStartTime = Get-Date
    Write-Host "Script started at: $($scriptStartTime)" -ForegroundColor Cyan
    Log-Message "Script started at: $($scriptStartTime)" "Info"


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
                Write-Host "Error processing site $siteUrl : $_" -ForegroundColor Red
                Log-Message "Error processing site $siteUrl : $_" "Error"
            }
            finally {
                $SPSite.Dispose()
            }
        }
        else {
            Write-Host "Could not access site $siteUrl" -ForegroundColor Yellow
            Log-Message "Could not access site $siteUrl" "Warning"
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
                Write-Host "Error processing web $webUrl : $_" -ForegroundColor Red
                Log-Message "Error processing web $webUrl : $_" "Error"
            }
            finally {
                $SPWeb.Dispose()
            }
        }
        else {
            Write-Host "Could not access web $webUrl" -ForegroundColor Yellow
            Log-Message "Could not access web $webUrl" "Warning"
        }
    }


    # Send the report via email
    if ($reportBody -ne "") {
        Send-Report
        Write-Host "Report sent successfully." -ForegroundColor Green
        Log-Message "Report sent successfully." "Info"
    }
    else {
        Write-Host "No data to report." -ForegroundColor Yellow
        Log-Message "No data to report." "Warning"
    }


    # Record the end time and calculate elapsed time
    $scriptEndTime = Get-Date
    Write-Host "Script ended at: $($scriptEndTime)" -ForegroundColor Cyan
    Log-Message "Script ended at: $($scriptEndTime)" "Info"
    $elapsedTime = $scriptEndTime - $scriptStartTime
    Write-Host "Total elapsed time: $($elapsedTime.Hours) hours, $($elapsedTime.Minutes) minutes, $($elapsedTime.Seconds) seconds." -ForegroundColor Cyan
    Log-Message "Total elapsed time: $($elapsedTime.Hours) hours, $($elapsedTime.Minutes) minutes, $($elapsedTime.Seconds) seconds." "Info"
}


# Run the main function
Main










