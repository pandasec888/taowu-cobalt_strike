function KillEvenlogService {
    # Kill Eventlog Service
    $EventlogSvchostPID = Get-WmiObject -Class win32_service -Filter "name = 'eventlog'" | select -exp ProcessId
    taskkill /F /PID $EventlogSvchostPID
}