#region Execution Details
    <#    5:30am - 10pm EST
    
        Runs every minute
	    >2 minutes = Alert
	    If it starts working > does an all clear.
	    15 good / bad every minute 

        If error - retry - email if fails. 
    #>
    #endregion


#region Functions: Send-Mail, Test-WebsiteContent, EventLogFunctions
    Import-Module E:\ScheduledTasks\Monitoring\Functions\Send-Mail.ps1
    Import-Module E:\ScheduledTasks\Monitoring\Functions\Test-WebSiteContent.ps1
    Import-Module E:\ScheduledTasks\Monitoring\Functions\EventLogFunctions.ps1
    #endregion

#region Redundancy Logic Variables
    $Primary = "App1"
    $Secondary = "App2"
    #endregion

#region Service IPs and Hostnames
    $LoadBalancers = 
        [PSCustomobject]@{IP = "1.2.3.4" ; Hostname = "LB-portal.domain.com"},
        [PSCustomObject]@{IP = "1.2.3.4" ; Hostname = "LB-intranet.domain.com"},
        [PSCustomObject]@{IP = "1.2.3.4" ; Hostname = "LB-mysite.domain.com"}

    $WebServers = 
        [PSCustomobject]@{IP = "1.2.3.4" ; Hostname = "web01"},
        [PSCustomobject]@{IP = "1.2.3.4" ; Hostname = "web02"},
        [PSCustomobject]@{IP = "1.2.3.4" ; Hostname = "web03"}

    $ReportServers = 
        [PSCustomobject]@{IP = "1.2.3.4" ; Hostname = "report01"}

    $AppServers = 
        [PSCustomobject]@{IP = "1.2.3.4" ; Hostname = "App1"},
        [PSCustomobject]@{IP = "1.2.3.4" ; Hostname = "App2"}
    #endregion Service IPs and Hostnames

#region Services and Fail Conditions
    $ServicesAndConditions = 
        [PSCustomObject]@{
            ServiceName = "Portal"
            Resources = @(
                $LoadBalancers | where {$_.hostname -like "*portal*"}
                $WebServers
            )
            BaseURL = "https://portal.domain.com"
            SearchString = '<div class="big-logo"><img src="/Assets/images/logo-large-portal-v3.png" border="0" /></div>'
        },
        [PSCustomObject]@{
            ServiceName = "Portal Search"
            Resources = @(
                $LoadBalancers | where {$_.hostname -like "*portal*"}
                $WebServers
            )
            BaseURL = "https://portal.domain.com/pages/results.aspx?k=test" 
            SearchString = '<a href="https://portal.domain.com/sites/sandbox">SharePoint Sandbox</a>'
        },
        [PSCustomObject]@{
            ServiceName = "Central Admin"
            Resources = @(
                $AppServers
            )
            BaseURL = "https://centadmin.domain.com/default.aspx" 
            SearchString = "brandingText:'PROD SP16 CENTRAL ADMIN'"
        },
        [PSCustomObject]@{
            ServiceName = "Intranet"
            Resources = @(
                $LoadBalancers | where {$_.hostname -like "*intranet*"}
                $WebServers
            )
            BaseURL = "https://intranet.domain.com" 
            SearchString = 'Home - Intranet Home/News'
        },
        [PSCustomObject]@{
            ServiceName = "MySites"
            Resources = @(
                $LoadBalancers | where {$_.hostname -like "*mysite*"}
                $WebServers
            )
            BaseURL = "https://mysite.domain.com" 
            SearchString = 'webTitle: "My Site Host"'
        },
        [PSCustomObject]@{
            ServiceName = "Power BI"
            Resources = @(
                [PSCustomObject]@{
                    IP = "1.2.3.4"
                    Hostname = "report01"
                }
            )
            BaseURL = "https://powerbi.domain.com/Reports/powerbi/SPTeam/pbi_test1" 
            SearchString = 'href="assets/styles.f63f5d9775e894a6ac4a.css"'
        }

        
    #Build a lookup hash for last errors for the purpose of throttling emails.  
        $LastHostErrorTimes = @{} ; ($LoadBalancers + $WebServers + $ReportServers + $AppServers) | ForEach-Object {$LastHostErrorTimes[$_.hostname] = $NULL}
    #endregion Services and Conditions

#Initiate log
    $Logsource = "Service-Monitoring"
    CreateLog $Logsource


#Endless loop
$N = 0
while($TRUE){
    $N++
    #Time condition check (only run during designated hours): 
        $CurrentTime     = Get-Date -Format HHmm
        $StartMonitoring = Get-Date -Hour 05 -Minute 30 -Format HHmm
        $EndMonitoring   = Get-Date -Hour 22 -Minute 0 -Format HHmm
    WriteLog -color Magenta -text "$($N) - $(Get-Date -Format HH:mm:ss) - $(Get-Date)"
    if($CurrentTime -in $StartMonitoring..$EndMonitoring){

        #region Redundancy Logic
            if($env:COMPUTERNAME -eq $Primary){
                WriteLog -color Green -text "  - $($N) - $(Get-Date -Format HH:mm:ss) - $(Get-Date -Format HH:mm:ss) - $($env:COMPUTERNAME) - PRIMARY"
                $RunScript = $True
            }
            elseif($env:COMPUTERNAME -eq $Secondary){
                try{
                    $RunScript = $FALSE
                    $primaryjob = invoke-command -computername $Primary -scriptblock {Get-ScheduledTask Service-Monitoring} | select *

                    if($primaryjob.State -ne 4){ 
                        #3 = ready
                        #4 = running
                    
                        <#RUN IT#>
                    }
                
                    WriteLog -color Gray -text "$($N) - $(Get-Date -Format HH:mm:ss) - $($env:COMPUTERNAME) - SECONDARY - Primary running. Skipping this cycle"
                }
                catch{
                    WriteLog -color Red -text "$($N) - $(Get-Date -Format HH:mm:ss) - $($env:computername) - SECONDARY - Primary failing."
                    $RunScript = $TRUE
                }
            }
            #endregion Redundancy Logic

    
        if($RunScript -eq $TRUE){
            WriteLog -color Yellow -text "  - $($N) - $(Get-Date -Format HH:mm:ss) - RUNNING TEST - $($ServicesAndConditions.servicename -join ", ")"
            #region Initial Check
            $Results = @()
                foreach($service in $ServicesAndConditions){
                    WriteLog -color Yellow -text "    - $($N) - $(Get-Date -Format HH:mm:ss) - $($service.ServiceName)"
                    $Results += Test-WebsiteContent `
                        -ServiceName  $service.ServiceName `
                        -WebServers   $service.Resources `
                        -BaseURL      $service.BaseURL `
                        -SearchString $service.SearchString
                }
            #endregion Initial Check
            
            #DEBUG
                #$Hold = $Results #DEBUG
                #$Results = $Hold #DEBUG
            #TEST ERROR CONDITION
                #$Results[1].Result = "False" ; $results[1].Error = "TEST ERROR"

            #Validates all Central Admin results if even 1 is up. 
                if($Results | where {($_.ServiceName -like "Central Admin") -and ($_.Result -like "True")})
                    { 
                        WriteLog -color Green -text "  - $($N) - $(Get-Date -Format HH:mm:ss) - Central Admin running successfully. Setting all instances to no-error."
                        $Results | where {$_.ServiceName -like "Central Admin"} | ForEach-Object {$_.Result = "True" ; $_.Error = "" } 
                    }

            #Runs through Error actions.
            if($Results | where {$_.Error}){
                
                foreach($failure in $Results | where {$_.Error} ){
                    WriteLog -color Red -text "  - $($N) - $(Get-Date -Format HH:mm:ss) - ERROR - "$failure.Hostname " - "$failure.Error
                    #Use a job to retry / notify if it happens again / with throttling to ensure no waking up to 1500 emails/txts... 
                    $ThrottleThreshold = 13 #Minutes - 2 (to account for sleep in loop) 
                    if( ($LastHostErrorTimes[$failure.hostname] -ge (Get-Date).AddMinutes($ThrottleThreshold)) `
                        -or ($NULL -eq $LastHostErrorTimes[$failure.Hostname]) ){
                        $LastHostErrorTimes[$failure.hostname] = Get-Date #Sets the last error per hostname so we only get 1 email per host error. 

                        
                        Start-Job `
                            -ScriptBlock {
                                Import-Module E:\ScheduledTasks\Monitoring\Functions\Send-Mail.ps1
                                Import-Module E:\ScheduledTasks\Monitoring\Functions\Test-WebSiteContent.ps1
            
                                Start-Sleep -Seconds 120 #Wait 2 minutes before trying again.

                                #Try again: 
                                    $LastTry = Test-WebsiteContent `
                                        -ServiceName  $args[0].ServiceName `
                                        -WebServers   $args[0].WebServers `
                                        -BaseURL      $args[0].BaseURL `
                                        -SearchString $args[0].SearchString
            
                                #Notify if still failing: 
                                    if($LastTry.error){
                                        Write-Output @{ "FAILED" = $args }
                                    } else {
                                        Write-Output @{"CLEARED" = $args}
                                    }
                                    
                            } -ArgumentList @( $failure )

                    } #if throttle end
                    else {WriteLog -color Red -text "  - $($N) - $(Get-Date -Format HH:mm:ss) - Error still occuring but not rechecking due to throttle"}
                    $EventLogID = 101 #ERRORS
                } #foreach failure in results end
        
            } #if results where error end.
            else { 
                WriteLog -color Green -text "  - $($N) - $(Get-Date -Format HH:mm:ss) - No errors detected. Looping." 
                $EventLogID = 1 #NO ERRORS
            }

            #Gets the previously run jobs to check for clear conditions / email if the state has cleared.
                $CompletedJobs = Get-Job | where {$_.State -eq "Completed"}

                foreach($CompletedJob in $CompletedJobs){
                    $JobOutput = Get-Job $CompletedJob.Id | Receive-Job
                    Remove-Job $CompletedJob.Id -Force
                    $ServiceAndHost = "$($JobOutput.values.ServiceName) - $($JobOutput.values.Hostname)"

                    if($JobOutput.Keys -eq "CLEARED"){
                        WriteLog -color Green -text "  - $($N) - $(Get-Date -Format HH:mm:ss) - $($ServiceAndHost) - CLEARED"

                        Send-Mail -subject ("CLEARED - " + $($ServiceAndHost)) -body "CLEARED ERROR - $($JobOutput.values.error)"
                        $LastHostErrorTimes[$JobOutput.Values.hostname] = $NULL
                        $EventLogID = 1 #NO ERRORS
                    } 
                    elseif($JobOutput.Keys -eq "FAILED"){
                        WriteLog -color Red -text " - $($N) - $(Get-Date -Format HH:mm:ss) - $($ServiceAndHost) - Recheck Failed - sending mail"

                        $Subject = "ERROR - " + $($ServiceAndHost)
                        Send-Mail -subject $Subject -body "$($Subject) - $($JobOutput.values.Error)"
                        $LastHostErrorTimes[$JobOutput.Values.hostname] = Get-Date
                        $EventLogID = 101 #ERRORS
                    }
                }
        } #if runscript -eq true end
        else {
            WriteLog -color Green -text "  - $($N) - $(Get-Date -Format HH:mm:ss) - Runscript Variable eq FALSE"
            $EventLogID = 1 #No errors.
        }

    }#if currenttime end
    else {
        WriteLog -color Green -text "  - $($N) - $(Get-Date -Format HH:mm:ss) - Outside monitoring hours."
        $EventLogID = 1 #No errors.
    }

#Write Logs
    if($EventLogID -eq 1){SaveLog -id 1 -txt "  - NO ERRORS"}
    elseif($EventLogID -eq 101){SaveLog -id 101 -error -txt "  - ERROR DETECTED"}

rv eventlogid,subject,joboutput,completedjobs,failure,result,results,msg -ea SilentlyContinue #DisposeVariables

Start-Sleep 120 #General wait between constant reattempts/time evaluations/etc.
}






    <#region Used to build a SearchString
        $BaseURL = "https://mysite.domain.com"
    
        $webclient = New-Object System.Net.WebClient
            $webClient.UseDefaultCredentials = $true
            $webclient.Headers.Add("host","mysite.domain.com")
            $page = $webclient.DownloadString($BaseURL)
            #View the $PAGE object and find something that is easily identified/should be there every single time. 
        
            #Tests the logic: 
                $pagepage.Contains('webTitle: "My Site Host"')
    #endregion Used to build a SearchString #>