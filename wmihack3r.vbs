Option Explicit
' Constants for WMI and registry operations
Const HKEY_LOCAL_MACHINE = &H80000002
Const wbemFlagReturnImmediately = &H10
Const wbemFlagForwardOnly = &H20
Const REG_KEY_PATH = "SOFTWARE\Classes\WMICommander"
Const REG_VALUE_NAME = "Output"
Const REG_FILE_VALUE_NAME = "FileData"
Const TIMEOUT = 3000 ' Timeout in milliseconds
Const TASK_NAME_LENGTH_MIN = 6
Const TASK_NAME_LENGTH_MAX = 12

' Global objects
Dim objWMIService, SubobjSWbemServices, regWMIService
Dim tempFilePath

' Main script logic
Sub Main()
    Dim args, mode, host, user, pass, command, getRes, localPath, remotePath
    Dim targetVersion, timeZone, execTime

    ' Parse command-line arguments
    Set args = WScript.Arguments
    If args.Count < 2 Or args.Count > 6 Then
        DisplayUsage
        WScript.Quit
    End If

    mode = LCase(args(0))
    host = args(1)
    Select Case mode
        Case "/cmd"
            If args.Count <> 6 Then DisplayUsage : WScript.Quit
            user = args(2) : pass = args(3) : command = args(4) : getRes = args(5)
        Case "/shell"
            If args.Count <> 4 Then DisplayUsage : WScript.Quit
            user = args(2) : pass = args(3)
        Case "/upload", "/download"
            If args.Count <> 6 Then DisplayUsage : WScript.Quit
            user = args(2) : pass = args(3) : localPath = args(4) : remotePath = args(5)
        Case Else
            DisplayUsage
            WScript.Quit
    End Select

    ' Connect to WMI services
    WScript.Echo "Connecting to " & host & "..."
    If Not ConnectWMI(host, user, pass) Then
        WScript.Echo "ERROR: Failed to connect to WMI. " & Err.Description
        WScript.Quit
    End If
    WScript.Echo "Connection established."

    ' Get target OS version
    targetVersion = GetOSVersion()
    WScript.Echo "Target OS major version: " & targetVersion

    ' Execute requested operation
    Select Case mode
        Case "/cmd"
            ExecuteCommand command, getRes, targetVersion
        Case "/shell"
            RunShell targetVersion
        Case "/upload"
            UploadFile localPath, remotePath
        Case "/download"
            DownloadFile localPath, remotePath
    End Select
End Sub

' Display usage instructions
Sub DisplayUsage()
    WScript.Echo "Usage:"
    WScript.Echo "  WMICommander.vbs /cmd host user pass command getres"
    WScript.Echo "  WMICommander.vbs /shell host user pass"
    WScript.Echo "  WMICommander.vbs /upload host user pass localpath remotepath"
    WScript.Echo "  WMICommander.vbs /download host user pass localpath remotepath"
    WScript.Echo ""
    WScript.Echo "Options:"
    WScript.Echo "  /cmd        Execute a single command"
    WScript.Echo "  /shell      Start an interactive shell"
    WScript.Echo "  /upload     Upload a file to the remote host"
    WScript.Echo "  /download   Download a file from the remote host"
    WScript.Echo "  host        Hostname or IP address"
    WScript.Echo "  user        Username (use '-' for default credentials)"
    WScript.Echo "  pass        Password (use '-' for default credentials)"
    WScript.Echo "  getres      1 to retrieve command output, 0 to execute without output"
End Sub

' Connect to WMI namespaces
Function ConnectWMI(host, user, pass)
    On Error Resume Next
    Dim objLocator
    Set objLocator = CreateObject("wbemscripting.swbemlocator")
    If user = "-" And pass = "-" Then
        Set objWMIService = objLocator.ConnectServer(host, "root/cimv2")
        Set SubobjSWbemServices = objLocator.ConnectServer(host, "root/subscription")
        Set regWMIService = objLocator.ConnectServer(host, "root/default")
    Else
        Set objWMIService = objLocator.ConnectServer(host, "root/cimv2", user, pass)
        Set SubobjSWbemServices = objLocator.ConnectServer(host, "root/subscription", user, pass)
        Set regWMIService = objLocator.ConnectServer(host, "root/default", user, pass)
    End If
    If Err.Number <> 0 Then
        ConnectWMI = False
    Else
        ConnectWMI = True
    End If
    On Error GoTo 0
End Function

' Get OS major version
Function GetOSVersion()
    Dim colItems, objItem, versionParts
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
    For Each objItem In colItems
        versionParts = Split(objItem.Version, ".")
        GetOSVersion = CInt(versionParts(0))
        Exit For
    Next
End Function

' Execute a single command
Sub ExecuteCommand(command, getRes, targetVersion)
    WScript.Echo "Executing: " & command
    Dim execTime, timeZone, taskName, outputFilePath, success
    If targetVersion < 6 Then
        tempFilePath = "C:\Windows\Temp\wmicmd_" & GenerateRandomString(TASK_NAME_LENGTH_MIN, TASK_NAME_LENGTH_MAX) & ".txt"
        execTime = GetExecutionTime(timeZone)
        If getRes = "1" Then
            ScheduleJobWithOutput command, tempFilePath, execTime, timeZone
            StoreOutputInRegistry tempFilePath
            DisplayOutput targetVersion
            DeleteRemoteFile tempFilePath
        Else
            ScheduleJobWithoutOutput command, execTime, timeZone
        End If
    Else
        taskName = GenerateRandomString(TASK_NAME_LENGTH_MIN, TASK_NAME_LENGTH_MAX)
        outputFilePath = "C:\Windows\Temp\wmicmd_" & taskName & ".txt"
        If getRes = "1" Then
            command = Replace(command, """", chr(34) & " & chr(34) & " & chr(34))
            success = CreateTaskWithOutput(command, outputFilePath, taskName)
            If success Then
                StoreOutputInRegistry outputFilePath
                DisplayOutput targetVersion
                DeleteRemoteFile outputFilePath
            Else
                WScript.Echo "ERROR: Command execution failed."
            End If
        Else
            command = Replace(command, """", chr(34) & " & chr(34) & " & chr(34))
            CreateTaskWithoutOutput command, taskName
        End If
    End If
End Sub

' Run interactive shell
Sub RunShell(targetVersion)
    Dim command
    WScript.Echo "WMICommander Shell (type 'exit' to quit)"
    Do
        WScript.StdOut.Write "CMD> "
        command = WScript.StdIn.ReadLine
        If LCase(Trim(command)) = "exit" Then Exit Do
        ExecuteCommand command, "1", targetVersion
    Loop
End Sub

' Upload a file
Sub UploadFile(localPath, remotePath)
    Dim binaryData
    binaryData = ReadBinaryFile(localPath)
    StoreBinaryInRegistry binaryData
    WriteFileFromRegistry remotePath
    WScript.Echo "File uploaded successfully."
End Sub

' Download a file
Sub DownloadFile(localPath, remotePath)
    ReadFileToRegistry remotePath
    Dim binaryData
    binaryData = RetrieveBinaryFromRegistry()
    WriteBinaryFile localPath, binaryData
    WScript.Echo "File downloaded successfully."
End Sub

' Get execution time and timezone
Function GetExecutionTime(ByRef timeZone)
    Dim colItems, objItem, execTime, tempTime
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_TimeZone", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
    For Each objItem In colItems
        timeZone = objItem.Bias
        If timeZone > 0 Then timeZone = "+" & timeZone
        Exit For
    Next
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_LocalTime", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
    For Each objItem In colItems
        execTime = FormatTime(objItem.Hour) & ":" & FormatTime(objItem.Minute) & ":" & FormatTime(objItem.Second)
        Exit For
    Next
    tempTime = DateAdd("s", 61, CDate(execTime))
    tempTime = Split(tempTime, ":")
    If CInt(tempTime(0)) < 10 Then tempTime(0) = "0" & tempTime(0)
    GetExecutionTime = tempTime(0) & tempTime(1)
End Function

' Format time to two digits
Function FormatTime(value)
    If value < 10 Then
        FormatTime = "0" & value
    Else
        FormatTime = value
    End If
End Function

' Schedule job for older systems with output
Sub ScheduleJobWithOutput(command, filePath, execTime, timeZone)
    Dim job, errCode, jobId
    execTime = "********" & execTime & "00.000000" & timeZone
    command = "c:\windows\system32\cmd.exe /c " & command & " > " & filePath
    Set job = objWMIService.Get("Win32_ScheduledJob")
    errCode = job.Create(command, execTime, True, , , True, jobId)
    If errCode <> 0 Then
        WScript.Echo "ERROR: Task creation failed."
        Exit Sub
    End If
    WScript.Echo "Task created. Waiting for execution..."
    WaitForFile filePath
End Sub

' Schedule job for older systems without output
Sub ScheduleJobWithoutOutput(command, execTime, timeZone)
    Dim job, errCode, jobId
    execTime = "********" & execTime & "00.000000" & timeZone
    command = "c:\windows\system32\cmd.exe /c " & command
    Set job = objWMIService.Get("Win32_ScheduledJob")
    errCode = job.Create(command, execTime, True, , , True, jobId)
    If errCode <> 0 Then
        WScript.Echo "ERROR: Task creation failed."
    Else
        WScript.Echo "Task created. Waiting for execution..."
    End If
End Sub

' Create task for newer systems with output
Function CreateTaskWithOutput(command, filePath, taskName)
    WScript.Echo "Task name: " & taskName
    Dim scriptText
    scriptText = "Const TriggerTypeDaily = 1" & vbCrLf & _
                 "Const ActionTypeExec = 0" & vbCrLf & _
                 "Set service = CreateObject(""Schedule.Service"")" & vbCrLf & _
                 "Call service.Connect" & vbCrLf & _
                 "Set rootFolder = service.GetFolder(""\"")" & vbCrLf & _
                 "Set taskDefinition = service.NewTask(0)" & vbCrLf & _
                 "Set regInfo = taskDefinition.RegistrationInfo" & vbCrLf & _
                 "regInfo.Description = ""Update""" & vbCrLf & _
                 "regInfo.Author = ""Microsoft""" & vbCrLf & _
                 "Set settings = taskDefinition.settings" & vbCrLf & _
                 "settings.Enabled = True" & vbCrLf & _
                 "settings.StartWhenAvailable = True" & vbCrLf & _
                 "settings.Hidden = False" & vbCrLf & _
                 "settings.DisallowStartIfOnBatteries = False" & vbCrLf & _
                 "Set triggers = taskDefinition.triggers" & vbCrLf & _
                 "Set trigger = triggers.Create(7)" & vbCrLf & _
                 "Set Action = taskDefinition.Actions.Create(ActionTypeExec)" & vbCrLf & _
                 "Action.Path = ""c:\windows\system32\cmd.exe""" & vbCrLf & _
                 "Action.arguments = chr(34) & ""/c " & command & " > " & filePath & """ & chr(34)" & vbCrLf & _
                 "Set objNet = CreateObject(""WScript.Network"")" & vbCrLf & _
                 "LoginUser = objNet.UserName" & vbCrLf & _
                 "If UCase(LoginUser) = ""SYSTEM"" Then" & vbCrLf & _
                 "    LoginUser = ""SYSTEM""" & vbCrLf & _
                 "Else" & vbCrLf & _
                 "    LoginUser = Empty" & vbCrLf & _
                 "End If" & vbCrLf & _
                 "Call rootFolder.RegisterTaskDefinition(""" & taskName & """, taskDefinition, 6, LoginUser, , 3)" & vbCrLf & _
                 "Call rootFolder.DeleteTask(""" & taskName & """, 0)"
    CreateWMIEventConsumer taskName, scriptText
    If WaitForFile(filePath) Then
        CreateTaskWithOutput = True
    Else
        CreateTaskWithOutput = False
    End If
End Function

' Create task for newer systems without output
Sub CreateTaskWithoutOutput(command, taskName)
    WScript.Echo "Task name: " & taskName
    Dim scriptText
    scriptText = "Const TriggerTypeDaily = 1" & vbCrLf & _
                 "Const ActionTypeExec = 0" & vbCrLf & _
                 "Set service = CreateObject(""Schedule.Service"")" & vbCrLf & _
                 "Call service.Connect" & vbCrLf & _
                 "Set rootFolder = service.GetFolder(""\"")" & vbCrLf & _
                 "Set taskDefinition = service.NewTask(0)" & vbCrLf & _
                 "Set regInfo = taskDefinition.RegistrationInfo" & vbCrLf & _
                 "regInfo.Description = ""Update""" & vbCrLf & _
                 "regInfo.Author = ""Microsoft""" & vbCrLf & _
                 "Set settings = taskDefinition.settings" & vbCrLf & _
                 "settings.Enabled = True" & vbCrLf & _
                 "settings.StartWhenAvailable = True" & vbCrLf & _
                 "settings.Hidden = False" & vbCrLf & _
                 "settings.DisallowStartIfOnBatteries = False" & vbCrLf & _
                 "Set triggers = taskDefinition.triggers" & vbCrLf & _
                 "Set trigger = triggers.Create(7)" & vbCrLf & _
                 "Set Action = taskDefinition.Actions.Create(ActionTypeExec)" & vbCrLf & _
                 "Action.Path = ""c:\windows\system32\cmd.exe""" & vbCrLf & _
                 "Action.arguments = chr(34) & ""/c " & command & """ & chr(34)" & vbCrLf & _
                 "Set objNet = CreateObject(""WScript.Network"")" & vbCrLf & _
                 "LoginUser = objNet.UserName" & vbCrLf & _
                 "If UCase(LoginUser) = ""SYSTEM"" Then" & vbCrLf & _
                 "    LoginUser = ""SYSTEM""" & vbCrLf & _
                 "Else" & vbCrLf & _
                 "    LoginUser = Empty" & vbCrLf & _
                 "End If" & vbCrLf & _
                 "Call rootFolder.RegisterTaskDefinition(""" & taskName & """, taskDefinition, 6, LoginUser, , 3)" & vbCrLf & _
                 "Call rootFolder.DeleteTask(""" & taskName & """, 0)"
    CreateWMIEventConsumer taskName, scriptText
End Sub

' Create WMI event consumer
Sub CreateWMIEventConsumer(name, scriptText)
    On Error Resume Next
    Dim asec, evtFlt, fcbnd, qstr, asecPath, evtFltPath, fcbndPath
    ' Create ActiveScriptEventConsumer
    Set asec = SubobjSWbemServices.Get("ActiveScriptEventConsumer").SpawnInstance_
    asec.Name = name
    asec.ScriptingEngine = "vbscript"
    asec.ScriptText = scriptText
    Set asecPath = asec.Put_()
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Failed to create ActiveScriptEventConsumer. " & Err.Description
        Exit Sub
    End If

    ' Create __EventFilter
    Set evtFlt = SubobjSWbemServices.Get("__EventFilter").SpawnInstance_
    evtFlt.Name = name & "_Filter"
    evtFlt.EventNameSpace = "root\cimv2"
    qstr = "SELECT * FROM __InstanceModificationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PerfFormattedData_PerfOS_System'"
    evtFlt.Query = qstr
    evtFlt.QueryLanguage = "wql"
    Set evtFltPath = evtFlt.Put_()
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Failed to create EventFilter. " & Err.Description
        SubobjSWbemServices.Delete asecPath.Path
        Exit Sub
    End If

    ' Create __FilterToConsumerBinding
    Set fcbnd = SubobjSWbemServices.Get("__FilterToConsumerBinding").SpawnInstance_
    fcbnd.Consumer = asecPath.Path
    fcbnd.Filter = evtFltPath.Path
    Set fcbndPath = fcbnd.Put_()
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Failed to create FilterToConsumerBinding. " & Err.Description
        SubobjSWbemServices.Delete asecPath.Path
        SubobjSWbemServices.Delete evtFltPath.Path
        Exit Sub
    End If

    ' Sleep to allow task execution
    WScript.Sleep 4000

    ' Cleanup
    On Error Resume Next
    SubobjSWbemServices.Delete fcbndPath.Path
    SubobjSWbemServices.Delete evtFltPath.Path
    SubobjSWbemServices.Delete asecPath.Path
    On Error GoTo 0
End Sub

' Wait for file to be created
Function WaitForFile(filePath)
    Dim replacedFile, query, colItems, objItem, done, retries, maxRetries
    replacedFile = Replace(filePath, "\", "\\")
    query = "SELECT * FROM CIM_DataFile WHERE name=" & chr(34) & replacedFile & chr(34)
    done = False
    retries = 0
    maxRetries = 10 ' 20 seconds total
    Do Until done Or retries >= maxRetries
        WScript.Sleep 2000
        Set colItems = objWMIService.ExecQuery(query, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
        For Each objItem In colItems
            WScript.Echo "File write successful."
            done = True
            Exit For
        Next
        retries = retries + 1
    Loop
    If Not done Then
        WScript.Echo "ERROR: File " & filePath & " was not created after " & maxRetries & " retries."
        WaitForFile = False
    Else
        WaitForFile = True
    End If
End Function

' Store command output in registry
Sub StoreOutputInRegistry(filePath)
    Dim scriptText
    scriptText = "Set ws = CreateObject(""wscript.shell"")" & vbCrLf & _
                 "Set fs = CreateObject(""scripting.filesystemobject"")" & vbCrLf & _
                 "If fs.FileExists(""" & filePath & """) Then" & vbCrLf & _
                 "    Set ts = fs.OpenTextFile(""" & filePath & """, 1)" & vbCrLf & _
                 "    content = ts.ReadAll" & vbCrLf & _
                 "    ts.Close" & vbCrLf & _
                 "    b64_content = Base64Encode(content, False)" & vbCrLf & _
                 "    path = ""HKEY_LOCAL_MACHINE\" & REG_KEY_PATH & "\""" & vbCrLf & _
                 "    ws.RegWrite path & ""Output"", b64_content" & vbCrLf & _
                 "Else" & vbCrLf & _
                 "    WScript.Echo ""ERROR: Output file not found.""" & vbCrLf & _
                 "End If" & vbCrLf & _
                 GetBase64EncodeFunction()
    CreateWMIEventConsumer "RegWriter", scriptText
End Sub

' Display command output
Sub DisplayOutput(targetVersion)
    Dim regValue
    If targetVersion < 6 Then
        regValue = GetRegistryStringValue(32)
    Else
        regValue = GetRegistryStringValue(64)
    End If
    If Not IsEmpty(regValue) Then
        On Error Resume Next
        WScript.Echo Base64Decode(regValue, False)
        If Err.Number <> 0 Then
            WScript.Echo "ERROR: Failed to decode Base64 output. " & Err.Description
        End If
        On Error GoTo 0
    Else
        WScript.Echo "ERROR: No output retrieved from registry."
    End If
End Sub

' Delete remote file
Sub DeleteRemoteFile(filePath)
    Dim replacedFile, query, colItems, objItem
    replacedFile = Replace(filePath, "\", "\\")
    query = "SELECT * FROM CIM_DataFile WHERE name=" & chr(34) & replacedFile & chr(34)
    Set colItems = objWMIService.ExecQuery(query, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
    For Each objItem In colItems
        objItem.Delete_
        Exit For
    Next
End Sub

' Read binary file
Function ReadBinaryFile(fileName)
    Dim buf, i, stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Mode = 3 : stream.Type = 1 : stream.Open : stream.LoadFromFile fileName
    ReDim buf(stream.Size - 1)
    For i = 0 To stream.Size - 1
        buf(i) = AscB(stream.Read(1))
    Next
    stream.Close
    ReadBinaryFile = buf
End Function

' Write binary file
Sub WriteBinaryFile(fileName, buf)
    Dim i, aBuf, size, bStream, tStream
    size = UBound(buf)
    ReDim aBuf(size \ 2)
    For i = 0 To size - 1 Step 2
        aBuf(i \ 2) = ChrW(buf(i + 1) * 256 + buf(i))
    Next
    If i = size Then aBuf(i \ 2) = ChrW(buf(i))
    aBuf = Join(aBuf, "")
    Set bStream = CreateObject("ADODB.Stream")
    bStream.Type = 1 : bStream.Open
    Set tStream = CreateObject("ADODB.Stream")
    tStream.Type = 2 : tStream.Open : tStream.WriteText aBuf
    tStream.Position = 2 : tStream.CopyTo bStream : tStream.Close
    bStream.SaveToFile fileName, 2 : bStream.Close
End Sub

' Store binary data in registry
Sub StoreBinaryInRegistry(binaryData)
    Dim reg
    Set reg = regWMIService.Get("StdRegProv")
    reg.CreateKey HKEY_LOCAL_MACHINE, REG_KEY_PATH
    Dim retCode
    retCode = reg.SetBinaryValue(HKEY_LOCAL_MACHINE, REG_KEY_PATH, REG_FILE_VALUE_NAME, binaryData)
    If retCode = 0 And Err.Number = 0 Then
        WScript.Echo "Binary data stored in registry."
    Else
        WScript.Echo "ERROR: Failed to store binary data. Code: " & retCode
    End If
End Sub

' Retrieve binary data from registry
Function RetrieveBinaryFromRegistry()
    Dim reg, binaryData
    Set reg = regWMIService.Get("StdRegProv")
    Dim retCode
    retCode = reg.GetBinaryValue(HKEY_LOCAL_MACHINE, REG_KEY_PATH, REG_FILE_VALUE_NAME, binaryData)
    If retCode = 0 And Err.Number = 0 Then
        RetrieveBinaryFromRegistry = binaryData

    Else
        WScript.Echo "ERROR: Failed to retrieve binary data. Code: " & retCode
        RetrieveBinaryFromRegistry = Empty
    End If
End Function

' Read file to registry
Sub ReadFileToRegistry(filePath)
    Dim scriptText
    scriptText = "arrData = ReadBinary(""" & filePath & """)" & vbCrLf & _
                 "Set objRegistry = GetObject(""winmgmts:{impersonationLevel=impPersonallyate}!\\.\root\default:StdRegProv"")" & vbCrLf & _
                 "objRegistry.CreateKey 2147483650, """ & REG_KEY_PATH & """" & vbCrLf & _
                 "retcode = objRegistry.SetBinaryValue(2147483650, """ & REG_KEY_PATH & """, """ & REG_FILE_VALUE_NAME & """, arrData)" & vbCrLf & _
                 "Function ReadBinary(FileName)" & vbCrLf & _
                 "  Dim Buf(), I" & vbCrLf & _
                 "  With CreateObject(""ADODB.Stream"")" & vbCrLf & _
                 "    .Mode = 3: .Type = 1: .Open: .LoadFromFile FileName" & vbCrLf & _
                 "    ReDim Buf(.Size - 1)" & vbCrLf & _
                 "    For I = 0 To .Size - 1: Buf(I) = AscB(.Read(1)): Next" & vbCrLf & _
                 "    .Close" & vbCrLf & _
                 "  End With" & vbCrLf & _
                 "  ReadBinary = Buf" & vbCrLf & _
                 "End Function"
    CreateWMIEventConsumer "FileReader", scriptText
    WScript.Echo "File read to registry."
End Sub

' Write file from registry
Sub WriteFileFromRegistry(filePath)
    Dim scriptText
    scriptText = "Set objRegistry = GetObject(""winmgmts:{impersonationLevel=impPersonallyate}!\\.\root\default:StdRegProv"")" & vbCrLf & _
                 "objRegistry.GetBinaryValue 2147483650, """ & REG_KEY_PATH & """, """ & REG_FILE_VALUE_NAME & """, strValue" & vbCrLf & _
                 "WriteBinary """ & filePath & """, strValue" & vbCrLf & _
                 "Sub WriteBinary(FileName, Buf)" & vbCrLf & _
                 "  Dim I, aBuf, Size, bStream" & vbCrLf & _
                 "  Size = UBound(Buf): ReDim aBuf(Size \ 2)" & vbCrLf & _
                 "  For I = 0 To Size - 1 Step 2" & vbCrLf & _
                 "      aBuf(I \ 2) = ChrW(Buf(I + 1) * 256 + Buf(I))" & vbCrLf & _
                 "  Next" & vbCrLf & _
                 "  If I = Size Then aBuf(I \ 2) = ChrW(Buf(I))" & vbCrLf & _
                 "  aBuf = Join(aBuf, """")" & vbCrLf & _
                 "  Set bStream = CreateObject(""ADODB.Stream"")" & vbCrLf & _
                 "  bStream.Type = 1: bStream.Open" & vbCrLf & _
                 "  With CreateObject(""ADODB.Stream"")" & vbCrLf & _
                 "    .Type = 2: .Open: .WriteText aBuf" & vbCrLf & _
                 "    .Position = 2: .CopyTo bStream: .Close" & vbCrLf & _
                 "  End With" & vbCrLf & _
                 "  bStream.SaveToFile FileName, 2: bStream.Close" & vbCrLf & _
                 "  Set bStream = Nothing" & vbCrLf & _
                 "End Sub"
    CreateWMIEventConsumer "FileWriter", scriptText
    WaitForFile filePath
    WScript.Echo "File written successfully."
End Sub

' Get registry string value
Function GetRegistryStringValue(architecture)
    Dim reg, ctx, inParams, outParams
    Set reg = regWMIService.Get("StdRegProv")
    Set ctx = CreateObject("WbemScripting.SWbemNamedValueSet")
    ctx.Add "__ProviderArchitecture", architecture
    ctx.Add "__RequiredArchitecture", True
    Set inParams = reg.Methods_("GetStringValue").InParameters
    inParams.hDefKey = HKEY_LOCAL_MACHINE
    inParams.sSubKeyName = REG_KEY_PATH
    inParams.sValueName = REG_VALUE_NAME
    Set outParams = reg.ExecMethod_("GetStringValue", inParams, , ctx)
    GetRegistryStringValue = outParams.sValue
End Function

' Base64 encode function
Function GetBase64EncodeFunction()
    GetBase64EncodeFunction = "Function Base64Encode(ByVal sText, ByVal fAsUtf16LE)" & vbCrLf & _
                              "    With CreateObject(""Msxml2.DOMDocument"").CreateElement(""aux"")" & vbCrLf & _
                              "        .DataType = ""bin.base64""" & vbCrLf & _
                              "        If fAsUtf16LE Then" & vbCrLf & _
                              "            .NodeTypedValue = StrToBytes(sText, ""utf-16le"", 2)" & vbCrLf & _
                              "        Else" & vbCrLf & _
                              "            .NodeTypedValue = StrToBytes(sText, ""utf-8"", 3)" & vbCrLf & _
                              "        End If" & vbCrLf & _
                              "        Base64Encode = .Text" & vbCrLf & _
                              "    End With" & vbCrLf & _
                              "End Function" & vbCrLf & _
                              "Function StrToBytes(ByVal sText, ByVal sTextEncoding, ByVal iBomByteCount)" & vbCrLf & _
                              "    With CreateObject(""ADODB.Stream"")" & vbCrLf & _
                              "        .Type = 2" & vbCrLf & _
                              "        .Charset = sTextEncoding" & vbCrLf & _
                              "        .Open" & vbCrLf & _
                              "        .WriteText sText" & vbCrLf & _
                              "        .Position = 0" & vbCrLf & _
                              "        .Type = 1" & vbCrLf & _
                              "        .Position = iBomByteCount" & vbCrLf & _
                              "        StrToBytes = .Read" & vbCrLf & _
                              "        .Close" & vbCrLf & _
                              "    End With" & vbCrLf & _
                              "End Function"
End Function

' Base64 decode
Function Base64Decode(sBase64EncodedText, fIsUtf16LE)
    Dim sTextEncoding
    If fIsUtf16LE Then sTextEncoding = "utf-16le" Else sTextEncoding = "utf-8"
    With CreateObject("Msxml2.DOMDocument").CreateElement("aux")
        .DataType = "bin.base64"
        .Text = sBase64EncodedText
        Base64Decode = BytesToStr(.NodeTypedValue, sTextEncoding)
    End With
End Function

' Convert bytes to string
Function BytesToStr(byteArray, sTextEncoding)
    If LCase(sTextEncoding) = "utf-16le" Then
        BytesToStr = CStr(byteArray)
    Else
        With CreateObject("ADODB.Stream")
            .Type = 1 : .Open : .Write byteArray
            .Position = 0 : .Type = 2 : .CharSet = sTextEncoding
            BytesToStr = .ReadText
            .Close
        End With
    End If
End Function

' Generate random string
Function GenerateRandomString(minLen, maxLen)
    Dim chars, i, length, result
    chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    length = RandomNumber(minLen, maxLen)
    result = ""
    For i = 1 To length
        result = result & Mid(chars, RandomNumber(1, Len(chars)), 1)
    Next
    GenerateRandomString = result
End Function

' Generate random number
Function RandomNumber(lowerBound, upperBound)
    Randomize
    RandomNumber = Int((upperBound - lowerBound + 1) * Rnd + lowerBound)
End Function

' Execute main script
Main