function LocalSessionManager {
    Try {
        Get-WinEvent -ListLog Security|out-null
    }
    Catch { return 'PowerShell Get-WinEvent cmdlet Error.' }
    Try {
        $SuccessResults=Get-WinEvent -LogName 'Microsoft-Windows-TerminalServices-LocalSessionManager/operational' -FilterXPath "*[System[(EventID=21) or (EventID=22) or (EventID=24) or (EventID=25)]]" -ErrorAction Stop
        $SuccessResults | Foreach {
            $entry = [xml]$_.ToXml()
            [array]$Output += New-Object PSObject -Property @{
                "TimeCreated" = $_.TimeCreated
                "EventID" = $entry.Event.System.EventID
                "EventRecordID" = $entry.Event.System.EventRecordID
                "User" = $entry.Event.UserData.EventXML.User
                "IpAddress" = $entry.Event.UserData.EventXML.Address
            }
        }
        $Output | Select-Object TimeCreated,EventID,EventRecordID,IpAddress
    }
    Catch { return 'Result: Null'}
}