Import-Module WebAdministration
$ActionTaken = $FALSE #Used to cycle IIS ONLY if needed.
 
$Config = [System.Web.Configuration.WebConfigurationManager]::OpenWebConfiguration("")
 
$httpruntime = $config.GetSection("system.web/httpRuntime")
Write-Host "EnableVersionHeader`n  - BEFORE: "$httpruntime.EnableVersionHeader
if($httpruntime.EnableVersionHeader -eq $TRUE){
	$httpruntime.EnableVersionHeader = $false
	$config.Save()
	$actiontaken = $TRUE
}
WRITE-HOST "  - AFTER: "$config.GetSection("system.web/httpRuntime").EnableVersionHeader
 
 
 

 #####Removes a specific header that is defined on each Web App, such as "MicrosoftSharePointTeamServices" which by default will tell the client which version of SharePoint is being used. Some cyber security divisons find this as a problem...
$headerToRemove = "MicrosoftSharePointTeamServices"
#Backup file:
	$WebConfigFile = Get-WebConfigFile
	$WebConfigBackupFile = ($WebConfigFile.DirectoryName + "\" + $WebConfigFile.name + "." + (Get-Date -Format yyyy-MM-dd))
	Copy-Item $WebConfigFile -Destination $WebConfigBackupFile -Verbose
 
$WebSites = Get-Website
 
foreach($Site in $WebSites){
$sitename = $Site.Name
Write-Output "Processing site: $sitename"
 
	try{$customheaders = Get-WebConfigurationProperty -PSPath "MACHINE/WEBROOT/APPHOST" -Filter "system.webServer/httpProtocol/customheaders" -Location $sitename -Name "." -ErrorAction SilentlyContinue}
	catch{Write-Host -Fore Red "Unable to get WebConfig for $sitename"}
	$property = $customheaders.Collection | where {$_.name -like $headerToRemove}
	if($property){
    	Write-Host -fore Yellow $headerToRemove " found on "$sitename " attempting to delete..."
    	$Result = Remove-WebConfigurationProperty -PSPath "MACHINE/WEBROOT/APPHOST" -Filter "system.webServer/httpProtocol/customheaders" -Location $sitename -AtElement @{name = $headerToRemove} -Name "." -Force -ErrorAction Inquire
    	Write-Host -fore Green "  - Success"
    	$ActionTaken = $TRUE
    	
	} else {
    	Write-Host -Fore Gray $headerToRemove " not found on" $sitename
	}
 
}

if($actiontaken){iisreset}


