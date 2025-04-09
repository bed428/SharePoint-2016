Function CreateLog {
        # EventLog - create source if missing
        if (!(Get-EventLog -LogName Application -Source $logSource -ErrorAction SilentlyContinue)) {
            New-EventLog -LogName Application -Source $logSource -ErrorAction SilentlyContinue | Out-Null
        }
    }

    Function WriteLog($text, $color) {
        $global:msg += "`n$text"
        if ($color) {
            Write-Host $text -Fore $color
        }
        else {
            Write-Output $text
        }
    }

    Function SaveLog($id, $txt, $error) {
        # EventLog
        if (!$skiplog) {
            if (!$error) {
                # Success
                $global:msg += $txt
                Write-EventLog -LogName Application -Source $logSource -EntryType Information -EventId $id -Message $global:msg
            }
            else {      
                # Error
                $global:msg += "ERROR`n"
                #$global:msg += $error.Message + "`n" + $error.ItemName
                Write-EventLog -LogName Application -Source $logSource -EntryType Warning -EventId $id -Message $global:msg
            }
        }
    }