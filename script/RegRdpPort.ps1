function RegRdpPort {
    $RegPath = "Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\"
    $RDPportValue = (Get-ItemProperty -Path $RegPath -ErrorAction Stop).PortNumber
    
    write-host "RDP-Tcp PortNumber: "$RDPportValue
}