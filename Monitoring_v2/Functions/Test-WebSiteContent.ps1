function Test-WebsiteContent {
        param (
        [Parameter(Mandatory=$true)]
        [PSObject[]]$WebServers,
        [string]$BaseURL,
        [string]$SearchString,
        [string]$ServiceName  # Added ServiceName parameter
        )

        # Extract the hostname from the BaseURL (only once):
        $BaseURLHostname = (New-Object Uri($BaseURL)).Host

        $Results = foreach ($WebServer in $WebServers) {
        try {
            $url = $BaseURL.Replace($BaseURLHostname, $WebServer.IP)

            $webClient = New-Object System.Net.WebClient
            $webClient.CachePolicy = [System.Net.Cache.HttpRequestCachePolicy]::BypassCache
            $webClient.UseDefaultCredentials = $true
            # Use the BaseURL's hostname in the Host header:
            $webClient.Headers.Add("host", $BaseURLHostname)

            $Searchpage = $webClient.DownloadString($url)
            $SearchTestResult = $Searchpage.Contains($SearchString)

            [PSCustomObject]@{
            ServiceName = $ServiceName        # Include ServiceName
            Hostname    = $WebServer.Hostname # Include Hostname
            IP          = $WebServer.IP
            Result      = $SearchTestResult
            WebServers  = $WebServers
            BaseURL     = $BaseURL
            SearchString= $SearchString
            Error       = $null # Indicate no error
            }
        }
        catch {
            [PSCustomObject]@{
            ServiceName = $ServiceName        # Include ServiceName
            Hostname    = $WebServer.Hostname # Include Hostname
            IP          = $WebServer.IP
            Result      = $false # Or $null, depending on your preference
            WebServers  = $WebServers
            BaseURL     = $BaseURL
            SearchString= $SearchString
            Error       = $_.Exception.Message
            }
        }
        } # foreach

        return $Results # Return the array of objects
    }