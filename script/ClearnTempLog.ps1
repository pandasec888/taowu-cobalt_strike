function ClearnTempLog {
    $TempLogPath = $env:temp
    write-host $TempLogPath
    Remove-Item $TempLogPath"\*" -Recurse -Force 2>&1 | Out-Null
}