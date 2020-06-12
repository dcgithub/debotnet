VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------
'Declaration  : Main Script class
'
'                CURRENTLY SUPPORTED:
'                - rm (Delete a file)
'                - rmdir (Delete a directory)
'                - Start (Start a file ONLY)
'                - Start()(Start a file with command prompt and wait for execution)
'                - Reg (Execute reg in command-line)
'                - Run (Execute commands)
'                - Try (Test reg queries in command-line) [TEST ONLY]
'                - Get (Download file with Wget in command-line)
'                - GetRaw (Download file RAW with Wget in command-line and paramter --content-disposition)
'                - UseRaw (Download file when hitting <Test script>) [TEST ONLY]
'                - OpenRaw (Open downloaded file when hitting <Test script> to preview script changes) [TEST ONLY]
'                - Url (Open a hyperlink in web per default web browser)
'                - Echo (Test Mode ONLY in command-line) [TEST ONLY]
'                - Msg (Displays a message box)
'                - TaskKill (Kill task with name or PID in command-line)
'                - FileExists (Check whether an file exists)
'---------------------------------------------------------------------------------

Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32.dll" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private m_IniFilename As String
Private m_Vars As New Collection

'---------------------------------------------------------------------------------
' Purpose :  Check if particular task is running
'---------------------------------------------------------------------------------

Private Declare Function CreateToolhelpSnapshot Lib _
  "kernel32" Alias "CreateToolhelp32Snapshot" ( _
  ByVal lFlgas As Long, ByVal lProcessID As Long) _
  As Long
 
Private Declare Function ProcessFirst Lib "kernel32" _
  Alias "Process32First" (ByVal hSnapshot As Long, _
  uProcess As PROCESSENTRY32) As Long
 
Private Declare Function ProcessNext Lib "kernel32" _
  Alias "Process32Next" (ByVal hSnapshot As Long, _
  uProcess As PROCESSENTRY32) As Long
 
Private Declare Sub CloseHandle Lib "kernel32" ( _
  ByVal hPass As Long)

Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Long = 260
 
Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szexeFile As String * MAX_PATH
End Type

'---------------------------------------------------------------------------------
' Purpose :  Check if particular task is running
'---------------------------------------------------------------------------------

Private Function IsEXERunning(ByVal sFilename As String) As Long
  
  Dim lSnapshot As Long
  Dim uProcess As PROCESSENTRY32
  Dim nResult As Long
  
  'Get snapshot of the current process
  lSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
  If lSnapshot <> 0 Then
    uProcess.dwSize = Len(uProcess)
    
    'Determine first process
    nResult = ProcessFirst(lSnapshot, uProcess)
    
    Do Until nResult = 0
      'Run through process list
      If InStr(LCase$(uProcess.szexeFile), LCase$(sFilename)) > 0 Then
        IsEXERunning = True
        Exit Do
      End If
      
      'Next process
      nResult = ProcessNext(lSnapshot, uProcess)
    Loop
    
    ' Close handle
    CloseHandle lSnapshot
  End If
End Function

Private Function FindFile(lzFilename As String) As Boolean

    If Trim(Len(lzFilename)) = 0 Then Exit Function
    FindFile = LenB(Dir(lzFilename)) <> 0
    
End Function

'Buffer set to 1024 char length

Private Function ExpandPath(ByVal LzPath As String) As String
Dim ret As Long
Dim nSize As Integer
Dim s_pos As Long
Dim sBuff As String

    'Space for buffer
    nSize = 1024
    'Create buffer
    sBuff = Space(nSize)
    'ExpandEnvironmentStrings
    ret = ExpandEnvironmentStrings(LzPath, sBuff, nSize)
    'Check for a good return.
    If (ret) Then
        s_pos = InStr(sBuff, Chr$(0))
    End If
    'Strip path
    If (s_pos) Then
        ExpandPath = Left(sBuff, s_pos - 1)
    End If
    'Clear up
    sBuff = vbNullString
    nSize = 0
    s_pos = 0
    ret = 0
    
End Function

Private Function SafeUBound(ByVal Arr) As Long

On Error GoTo TErr:
    'Return Safe Array Size without error.
    SafeUBound = UBound(Arr)
    Exit Function
TErr:
    SafeUBound = -1
End Function

'---------------------------------------------------------------------------------
'Declaration  : Used to read a value from a given keyname
'---------------------------------------------------------------------------------

Private Function INIReadKey(ByVal Selection As String, ByVal KeyName As String, Optional Default = "") As String

Dim ret As Long
Dim sBuff As String
    sBuff = Space(2048)
    ret = GetPrivateProfileString(Selection, KeyName, "{ERR}", sBuff, 2048, m_IniFilename)
    
    If (ret) Then
        'Return the keyvalue
        sBuff = Left(sBuff, ret)
        If (sBuff = "{ERR}") Then
            INIReadKey = Default
        Else
            INIReadKey = sBuff
        End If
    End If
    
    sBuff = ""
End Function

'---------------------------------------------------------------------------------
'Declaration  : Used to collect selections name of an INI File
'---------------------------------------------------------------------------------

Private Function GetSelectionNames() As Collection

Dim Col As New Collection
Dim fp As Long
Dim sLine As String
    
    fp = FreeFile
    Open m_IniFilename For Input As #fp
        Do Until EOF(fp)
            Line Input #fp, sLine
            sLine = Trim(sLine)
            If Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
                Col.add Mid(sLine, 2, Len(sLine) - 2)
            End If
        Loop
    Close #fp
    
    Set GetSelectionNames = Col
    Set Col = Nothing
    
End Function

'---------------------------------------------------------------------------------
'Declaration  : Collects all the keys in an INI's Selection
'---------------------------------------------------------------------------------

Private Function GetIniSelection(ByVal Selection As String) As Collection

Dim ret As Long
Dim count As Integer
Dim sBuff As String
Dim vLst() As String
Dim Tmp As String
Dim Col As New Collection

    sBuff = Space(8192)
    ret = GetPrivateProfileString(Selection, 0&, "Err", sBuff, 8192, m_IniFilename)
    
    If (ret > 0) Then
        Tmp = Left(sBuff, ret)
        vLst = Split(Tmp, Chr(0))
        
        'Loop though all the keynames
        For count = 0 To SafeUBound(vLst)
            If Len(vLst(count)) <> 0 Then
                'Add the found keyname
                Col.add vLst(count)
            End If
        Next count
    End If
    
    'Return the collection
    Set GetIniSelection = Col
    Set Col = Nothing
    Erase vLst
    sBuff = ""
    Tmp = ""
    count = 0
End Function

Public Property Get FileName() As String
    FileName = m_IniFilename
End Property

Public Property Let FileName(ByVal NewFileName As String)
    m_IniFilename = NewFileName
    'Phase script maybe!
    'Call PhasePlugUndo
End Property

' PIPELINE!
Private Sub DoVariables(ByVal Selection As String)

Dim iVals As New Collection
Dim sVal As String

    Set iVals = GetIniSelection(Selection)
    
    If (iVals.count > 0) Then
        For Each VarName In iVals
            sVal = Trim(INIReadKey(Selection, VarName))

            If Len(sVal) <> 0 Then
                'Store the variable name so we can free it latter
                m_Vars.add VarName
                'Add the variable to the current process
                SetEnvironmentVariable VarName, ExpandPath(sVal)
            End If
        Next VarName
    End If
    
    'Clear Up
    Set iVals = Nothing
    sVal = vbNullString
    
End Sub

'---------------------------------------------------------------------------------
'Purpose : Do the TEST
'---------------------------------------------------------------------------------

Public Sub PhasePlugTest()
Dim ColA As New Collection
Dim Item As Variant

    'Main scripts Loop
    'Load INI Selections
    Set ColA = GetSelectionNames
    
    'Loop through all Selections
    For Each Item In ColA
        Select Case UCase(Item)
         Case "VARIABLES"
                Call DoVariables(Item)
            Case "CODE"
                Call DoTestOps(Item)
        End Select
    Next Item
    
    'Free the variables data
    For Each Item In m_Vars
        SetEnvironmentVariable Item, vbNullString
    Next Item
    
    'Clean Up
    Set m_Vars = Nothing
    Set ColA = Nothing
    Item = ""
    
End Sub

'---------------------------------------------------------------------------------
'Purpose : Do the RUN
'---------------------------------------------------------------------------------

Public Sub PhasePlugRun()

Dim ColA As New Collection
Dim Item As Variant

    'Main Scripts Loop
    'Load INI Selections.
    Set ColA = GetSelectionNames
    
    'Loop tho all the selections.
    For Each Item In ColA
        Select Case UCase(Item)
            Case "VARIABLES"
                Call DoVariables(Item)
            Case "CODE"
                Call DoRunOps(Item)
        End Select
    Next Item
    
    'Free the variables data
    For Each Item In m_Vars
        SetEnvironmentVariable Item, vbNullString
    Next Item
    
    'Clean Up
    Set m_Vars = Nothing
    Set ColA = Nothing
    Item = ""
End Sub

'---------------------------------------------------------------------------------
'Purpose : Do the UNDO
'---------------------------------------------------------------------------------

Public Sub PhasePlugUndo()
Dim ColA As New Collection
Dim Item As Variant

    'Main Scripts Loop
    'Load INI Selections
    Set ColA = GetSelectionNames
    
    'Loop through all Selections
    For Each Item In ColA
        Select Case UCase(Item)
        
         Case "VARIABLES"
                Call DoVariables(Item)
            Case "UNDO"
                Call DoUndoOps(Item)
        End Select
           
    Next Item
    
    
    'Free the variables data
    For Each Item In m_Vars
        SetEnvironmentVariable Item, vbNullString
    Next Item
    
    'Clean Up
    Set m_Vars = Nothing
    Set ColA = Nothing
    Item = ""
    
End Sub

'---------------------------------------------------------------------------------
'Purpose : // Run TEST Operations \\
'---------------------------------------------------------------------------------

Public Function DoTestOps(ByVal Selection As String)

Dim iVals As New Collection
Dim sCmd As String
Dim sFilename As String
Dim vLst() As String

    Set iVals = GetIniSelection(Selection)
    
    If (iVals.count > 0) Then
        For Each fileop In iVals
            vLst = Split(Trim(INIReadKey(Selection, fileop)), ",", , vbBinaryCompare)
            sCmd = Trim(vLst(0))
            
            If SafeUBound(vLst) >= 1 Then
            
                sFilename = ExpandPath(vLst(1))
        
      
                Select Case sCmd
                
'---------------------------------------------------------------------------------
'Purpose : Delete file
'---------------------------------------------------------------------------------
                  Case "rm"
                    
                        Call WriteActions(frmMain.lstDS.Text)
                        Call WriteActions(sFilename)
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                        
'---------------------------------------------------------------------------------
'Purpose : Delete folder
'---------------------------------------------------------------------------------
                Case "rmdir"
            
                        Call WriteActions(frmMain.lstDS.Text)
                        Call WriteActions(sFilename)
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                        
'---------------------------------------------------------------------------------
'Purpose : Run file
'---------------------------------------------------------------------------------
                    Case "Start"
                    
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("[Start] " & sFilename)
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                            
'---------------------------------------------------------------------------------
'Purpose : Start (executables)
'---------------------------------------------------------------------------------
                    Case "Start()"
                    
                        Call WriteActions(frmMain.lstDS.Text)
                        Call WriteActions("[Start(Wait)] " & sFilename)
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                
'---------------------------------------------------------------------------------
'Purpose : Open a hyperlink
'---------------------------------------------------------------------------------
                    Case "Url"
                    
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("[Url] " & sFilename)
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                            
'---------------------------------------------------------------------------------
'Purpose : Run commands with reg.exe in command-line
'---------------------------------------------------------------------------------
                    Case "Reg"
         
                        If vLst(UBound(vLst)) = UCase("STDOUT") Then 'Show optional Command prompt window
                            
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("//Here the system changes which would be applied")
                            Call WriteActions("[Reg] " & sFilename)
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                
                        ElseIf vLst(UBound(vLst)) = UCase("SHOWCLI") Then ' Show optional Command prompt window
                         
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("//Results are opened in command prompt")
                            RetVal = ShellWait("cmd.exe /K" & sFilename, True)
    
                        Else
                               
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("//Here the system changes which would be applied")
                            Call WriteActions("[Reg] " & sFilename)
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    
                                   
                        End If
                        
 
'---------------------------------------------------------------------------------
'Purpose : Run a Test (query) ONLY of reg.exe with command-line and capture output
'---------------------------------------------------------------------------------

                    Case "Try"
                    
                      'x64 (Using the virtual "Sysnative" folder to access 64-bit reg.exe from 32-bit code)
                        'more Info here http://scottiestech.info/2009/06/20/calling-a-32-bit-system-command-from-a-script-in-x64-windows
                         If DirExists(Environ("WinDir") & "\SysWOW64") Then
                        
                             If vLst(UBound(vLst)) = UCase("STDOUT") Then 'Show optional Command prompt window
                             
                                 Call WriteActions(frmMain.lstDS.Text)
                                 Call WriteActions("[Test run] " & sFilename)
                                 Call WriteActions("//Here the results of testing")
                                 Call WriteActions(GetCommandOutput(Environ("SystemRoot") & "\sysnative\reg.exe" & " " & sFilename, True, True, True, 0))
                                 Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                 
                            ElseIf vLst(UBound(vLst)) = UCase("SHOWCLI") Then ' Show optional Command prompt window
                         
                                Call WriteActions(frmMain.lstDS.Text)
                                Call WriteActions("//Results are opened in command prompt")
                                RetVal = ShellWait(Environ("SystemRoot") & "\sysnative\reg.exe" & " " & sFilename, True)
                                Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                 
                            Else 'Silent
                                   
                                Call WriteActions(frmMain.lstDS.Text)
                                Call WriteActions("[Test run] " & Environ("SystemRoot") & "\sysnative\reg.exe" & " " & sFilename)
                                Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        
                            End If
                            
                    Else
                        
                        'x86-bit
                        
                            If vLst(UBound(vLst)) = UCase("STDOUT") Then 'Show optional Command prompt window
                                 
                                Call WriteActions(frmMain.lstDS.Text)
                                Call WriteActions("[Test run] " & sFilename)
                                Call WriteActions("//Here the results of testing")
                                Call WriteActions(GetCommandOutput("reg.exe" & " " & sFilename, True, True, True, 0))
                                Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                     
                            ElseIf vLst(UBound(vLst)) = UCase("SHOWCLI") Then ' Show optional Command prompt window
                             
                                Call WriteActions(frmMain.lstDS.Text)
                                Call WriteActions("//Results are opened in command prompt")
                                RetVal = ShellWait("cmd.exe /K reg.exe" & " " & sFilename, True)
                                Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                    
                            Else 'Silent
                                       
                                Call WriteActions(frmMain.lstDS.Text)
                                Call WriteActions("[Test run] " & "reg.exe" & " " & sFilename)
                                Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
            
                            End If
                            
                        
                        End If
                        
'---------------------------------------------------------------------------------
'Purpose : Run commands in command-line
'---------------------------------------------------------------------------------
                    Case "Run"
         
                        If vLst(UBound(vLst)) = UCase("STDOUT") Then 'Show optional Command prompt window
                            
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("//Here the system changes which would be applied")
                            Call WriteActions("[Run]" & " " & sFilename)
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                
                        ElseIf vLst(UBound(vLst)) = UCase("SHOWCLI") Then ' Show optional Command prompt window
                            
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("//Here the system changes which would be applied")
                            Call WriteActions("[Run]" & " " & sFilename)
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    
                        Else
                               
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("//Here the system changes which would be applied")
                            Call WriteActions("[Run] " & sFilename)
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        
                    End If
                        
'---------------------------------------------------------------------------------
'Purpose : Download files with Wget
'---------------------------------------------------------------------------------
                    Case "Get"
                
                        Call WriteActions(frmMain.lstDS.Text)
                        Call WriteActions("[Download " & GetLastSlash(sFilename) & ", " & sFilename & "]")
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                        
'---------------------------------------------------------------------------------
'Purpose : Show ONLY which files would be downloaded RAW with Wget
'---------------------------------------------------------------------------------
                    Case "GetRaw"
                
                        Call WriteActions(frmMain.lstDS.Text)
                        Call WriteActions("[Download " & sFilename & "]")
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

'---------------------------------------------------------------------------------
'Purpose : Download files RAW with Wget and use this optionally for execution in [Run] section
'---------------------------------------------------------------------------------
                    Case "UseRaw"
                                     
                        If vLst(UBound(vLst)) = UCase("DEFAULT") Then 'Download to directory set in frmMain.txtOutputDir.Text
                                        
                           Call WriteActions(frmMain.lstDS.Text)
                           Call WriteActions(GetCommandOutput("" & Chr(34) & frmMain.txtWgetPath.Text & """" & " " & Chr(34) & sFilename & "?raw=true" & """" & " " & "--show-progress --progress=bar:force --no-hsts --no-check-certificate --content-disposition -N -P" & " " & Chr(34) & frmMain.txtOutputDir.Text & """", True, True, True, 0))
                           Call WriteActions("Downloading " & sFilename & " completed")
                           Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        
    
                       Else
    
                           Call WriteActions(frmMain.lstDS.Text)
                           Call WriteActions(GetCommandOutput("" & Chr(34) & frmMain.txtWgetPath.Text & """" & " " & Chr(34) & sFilename & "?raw=true" & """" & " " & "--show-progress --progress=bar:force --no-hsts --no-check-certificate --content-disposition -N -P" & " " & Chr(34) & vLst(UBound(vLst)) & """", True, True, True, 0))
                           Call WriteActions("Downloading " & sFilename & " completed")
                           Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        
                      End If
                        
'---------------------------------------------------------------------------------
'Purpose : Open files RAW as plain text file
'---------------------------------------------------------------------------------
                    Case "OpenRaw"
                    
                         Call WriteActions(frmMain.lstDS.Text)
                         Call WriteActions("[Content of script file " & sFilename & " has been opened. Please review the changes.]")
                         
                         DocumentOpen sFilename
                    
                         Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                     
'---------------------------------------------------------------------------------
'Purpose : Show command-line ONLY
'---------------------------------------------------------------------------------
                    Case "Echo"
                    
                        Call WriteActions(frmMain.lstDS.Text)
                        Call WriteActions("//Here the system changes which would be applied")
                        Call WriteActions("[Run] " & sFilename)
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                
'---------------------------------------------------------------------------------
'Purpose : Check if file exisits
'---------------------------------------------------------------------------------
                    Case "FileExists"
                    
                        If FileExists(sFilename) = True Then
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("//Check whether file exists")
                            Call WriteActions("[Check] " & sFilename & " <true>")
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                            
                        Else
                            
                            Call WriteActions("//Check whether file exists")
                            Call WriteActions("[Check] " & sFilename & " <false>")
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                            
                        End If
                        
           
                 End Select
                
            End If
   
        Next fileop
    End If

    'Clear Up
    Set iVals = Nothing
    Erase vLst
    sVal = vbNullString
    sFilename = vbNullString
    sCmd = vbNullString

End Function

'---------------------------------------------------------------------------------
'Purpose : Run Script Operations
'---------------------------------------------------------------------------------

Private Function DoRunOps(ByVal Selection As String)

Dim iVals As New Collection
Dim sCmd As String
Dim sFilename As String
Dim ClonePath As String
Dim vLst() As String
                
Set iVals = GetIniSelection(Selection)
    
    If (iVals.count > 0) Then
        For Each fileop In iVals
            vLst = Split(Trim(INIReadKey(Selection, fileop)), ",", , vbBinaryCompare)
            sCmd = Trim(vLst(0))
            
            If SafeUBound(vLst) >= 1 Then
            
                sFilename = ExpandPath(vLst(1))
                
                Select Case sCmd
                         
'---------------------------------------------------------------------------------
'Purpose : Delete file
'---------------------------------------------------------------------------------
                Case "rm"
                    
                    If vLst(UBound(vLst)) = UCase("RECURSE") Then
                                
                        Call WriteActions("[Delete file(s) /recursively] " & sFilename)
                        RetVal = ShellWait("cmd.exe /C" & "del /s /q " & Chr(34) & sFilename & Chr(34), False)
                        Call WriteActions("[Delete file /recursively] " & sFilename & " <done>")
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                    Else
                        Call WriteActions("[Delete file] " & sFilename)
                        RetVal = ShellWait("cmd.exe /c" & "DEL /F /Q " & Chr(34) & sFilename & Chr(34), False)
                        Call WriteActions("[Delete file] " & sFilename & " <done>")
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                    End If
                        
'---------------------------------------------------------------------------------
'Purpose : Delete folder
'---------------------------------------------------------------------------------
                Case "rmdir"

                    If vLst(UBound(vLst)) = UCase("RECURSE") Then
                        
                        Call WriteActions("[Delete directory /recursively] " & sFilename)
                        RetVal = ShellWait("cmd.exe /C" & "rmdir /s /q " & Chr(34) & sFilename & Chr(34), False)
                        Call WriteActions("[Delete directory /recursively] " & sFilename & " <done>")
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                    Else
                        Call WriteActions("[Delete directory] " & sFilename)
                        RetVal = ShellWait("cmd.exe /C" & "DEL /F /Q " & Chr(34) & sFilename & Chr(34), False)
                        Call WriteActions("[Delete directory] " & sFilename & " <done>")
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                    End If

'---------------------------------------------------------------------------------
'Purpose : Start file ONLY
'---------------------------------------------------------------------------------
                    Case "Start"
                    
                        Call WriteActions(frmMain.lstDS.Text)
                        Call WriteActions("[Run] " & sFilename & " <processing>")
                                              
                        Call ShellExecute(hwnd, "Open", _
                                 sFilename, "", "", 1)
                             
                        Call WriteActions("[Run] " & sFilename & " <done>")
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                      
'---------------------------------------------------------------------------------
'Purpose : Start and wait for execution with command-line
'---------------------------------------------------------------------------------
                    Case "Start()"
                    
                        If DirExists(Environ("WinDir") & "\SysWOW64") Then  'x64 (Using the virtual "Sysnative" folder to access 64-bit cmd.exe from 32-bit code)
                            
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("[Start(Wait)] " & sFilename & " <processing>")
                                                  
                            RetVal = ShellWait(Environ("SystemRoot") & "\sysnative\cmd.exe /C" & Chr(34) & sFilename & Chr(34), False)
                                 
                            Call WriteActions("[Start(Wait)] " & sFilename & " <done>")
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    
                        Else 'x86-bit
                         
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("[Start(Wait)] " & sFilename & " <processing>")

                            RetVal = ShellWait("cmd.exe /C" & Chr(34) & sFilename & Chr(34), False)
                                   
                            Call WriteActions("[Start(Wait)] " & sFilename & " <done>")
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                            
                        End If
                                                                                                         
'---------------------------------------------------------------------------------
'Purpose : Open a hyperlink
'---------------------------------------------------------------------------------
                    Case "Url"
                    
                        Call WriteActions(frmMain.lstDS.Text)
                        Call WriteActions("[Url] " & sFilename & " <processing>")
                                              
                        Call ShellExecute(hwnd, "Open", _
                                 sFilename, "", "", 1)
                             
                        Call WriteActions("[Url] " & sFilename & " <done>")
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                      
                                                               
'---------------------------------------------------------------------------------
'Purpose : Run commands with reg.exe in command-line
'---------------------------------------------------------------------------------
                    Case "Reg"
                    
                     'x64 (Using the virtual "Sysnative" folder to access 64-bit reg.exe from 32-bit code)
                        'more Info here http://scottiestech.info/2009/06/20/calling-a-32-bit-system-command-from-a-script-in-x64-windows
                         If DirExists(Environ("WinDir") & "\SysWOW64") Then
                    
                            If vLst(UBound(vLst)) = UCase("STDOUT") Then 'Capture output
                            
                                   Call WriteActions(frmMain.lstDS.Text)
                                   Call WriteActions("//Here the results of running")
                                   Call WriteActions("[Reg] " & sFilename)
                                   Call WriteActions(GetCommandOutput(Environ("SystemRoot") & "\sysnative\reg.exe" & " " & sFilename, True, True, True, 0))
                                   Call WriteActions("[Reg] " & frmMain.lstDS.Text & " <done>")
                                   Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                               
                            ElseIf vLst(UBound(vLst)) = UCase("SHOWCLI") Then ' Show optional Command prompt window
                            
                                   Call WriteActions(frmMain.lstDS.Text)
                                   Call WriteActions("//Results are opened in command prompt")
                                   RetVal = ShellWait(Environ("SystemRoot") & "\sysnative\reg.exe /C" & " " & sFilename, True)
                                   Call WriteActions("[Reg] " & frmMain.lstDS.Text & " <done>")
                                   Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                               
                              Else  'Run silent
                              
                                   Call WriteActions(frmMain.lstDS.Text)
                                   Call WriteActions("[Reg] " & sFilename & " <processing>")
                                   RetVal = ShellWait(Environ("SystemRoot") & "\sysnative\reg.exe /C" & " " & sFilename, False)
                                   Call WriteActions("[Reg] " & frmMain.lstDS.Text & " <done>")
                                   Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                
                                  
                              End If
                              
                        Else
                        
                         'x86
                    
                            If vLst(UBound(vLst)) = UCase("STDOUT") Then 'Capture output
                            
                                   Call WriteActions(frmMain.lstDS.Text)
                                   Call WriteActions("//Here the results of running")
                                   Call WriteActions("[Reg] " & sFilename)
                                   Call WriteActions(GetCommandOutput("reg.exe" & " " & sFilename, True, True, True, 0))
                                   Call WriteActions("[Reg] " & frmMain.lstDS.Text & " <done>")
                                   Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                               
                            ElseIf vLst(UBound(vLst)) = UCase("SHOWCLI") Then ' Show optional Command prompt window
                            
                                   Call WriteActions(frmMain.lstDS.Text)
                                   Call WriteActions("//Results are opened in command prompt")
                                   RetVal = ShellWait("cmd /K reg.exe" & " " & sFilename, True)
                                   Call WriteActions("[Reg] " & frmMain.lstDS.Text & " <done>")
                                   Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                               
                              Else  'Run silent
                              
                                   Call WriteActions(frmMain.lstDS.Text)
                                   Call WriteActions("[Reg] " & sFilename & " <processing>")
                                   RetVal = ShellWait("reg.exe" & " " & sFilename, False)
                                   Call WriteActions("[Reg] " & frmMain.lstDS.Text & " <done>")
                                   Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                
                                  
                              End If
                    
                    End If
                    
'---------------------------------------------------------------------------------
'Purpose : Run commands in command-line
'---------------------------------------------------------------------------------
                    Case "Run"
        
                        If vLst(UBound(vLst)) = UCase("STDOUT") Then 'Capture output
                            
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("//Here the results of running")
                            Call WriteActions("[Run] " & sFilename & " <processing>")
                            Call WriteActions(GetCommandOutput(sFilename, True, True, True, 0))
                            Call WriteActions("[Run] " & frmMain.lstDS.Text & " <done>")
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                      
                        ElseIf vLst(UBound(vLst)) = UCase("SHOWCLI") Then ' Show optional Command prompt window
                                   
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("//Results are opened in command prompt")
                            RetVal = ShellWait("cmd.exe /K" & sFilename, True)
                            Call WriteActions("[Run] " & frmMain.lstDS.Text & " <done>")
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                      
                        Else  'Run silent
                                     
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("[Run] " & sFilename & " <processing>")
                            RetVal = ShellWait("cmd.exe /C" & sFilename, False)
                            Call WriteActions("[Run] " & frmMain.lstDS.Text & " <done>")
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                               
                        End If
                 
'---------------------------------------------------------------------------------
'Purpose : Download files with Wget
'---------------------------------------------------------------------------------
                  Case "Get"
      
                     If vLst(UBound(vLst)) = UCase("DEFAULT") Then 'Download to directory set in frmMain.txtOutputDir.Text
                                    
                       Call WriteActions(frmMain.lstDS.Text)
                       Call WriteActions(GetCommandOutput("" & Chr(34) & frmMain.txtWgetPath.Text & """" & " " & sFilename & " " & frmMain.txtWgetParam.Text & " " & Chr(34) & frmMain.txtOutputDir.Text & "", True, True, True, 0))
                       Call WriteActions("Downloading " & sFilename & " completed")
                       Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    
              
                   Else

                       Call WriteActions(frmMain.lstDS.Text)
                       Call WriteActions(GetCommandOutput("" & Chr(34) & frmMain.txtWgetPath.Text & """" & " " & sFilename & " " & frmMain.txtWgetParam.Text & " " & Chr(34) & vLst(UBound(vLst)) & "", True, True, True, 0))
                       Call WriteActions("Downloading " & sFilename & " completed")
                       Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    
                    End If

'---------------------------------------------------------------------------------
'Purpose : Download files RAW with Wget paramter --content-disposition
'---------------------------------------------------------------------------------
                Case "GetRaw"
                
                    'Check if script file already exisits (this could be used as customized version)
                    If FileExists((vLst(UBound(vLst)) & "\" & Mid$(sFilename, InStrRev(sFilename, "/") + 1))) Then
                    
                        Call WriteActions("[Download skipped]")
                        Call WriteActions("Script already present in " & (vLst(UBound(vLst)) & "\" & Mid$(sFilename, InStrRev(sFilename, "/") + 1)))
                        
                    Else
                        If vLst(UBound(vLst)) = UCase("DEFAULT") Then 'Download to directory set in frmMain.txtOutputDir.Text
                                        
             
                           Call WriteActions(frmMain.lstDS.Text)
                           Call WriteActions(GetCommandOutput("" & Chr(34) & frmMain.txtWgetPath.Text & """" & " " & Chr(34) & sFilename & "?raw=true" & """" & " " & "--show-progress --progress=bar:force --no-hsts --no-check-certificate --content-disposition -N -P" & " " & Chr(34) & frmMain.txtOutputDir.Text & """", True, True, True, 0))
                           Call WriteActions("Downloading " & sFilename & " completed")
                           Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        
    
                       Else
    
                           Call WriteActions(frmMain.lstDS.Text)
                           Call WriteActions(GetCommandOutput("" & Chr(34) & frmMain.txtWgetPath.Text & """" & " " & Chr(34) & sFilename & "?raw=true" & """" & " " & "--show-progress --progress=bar:force --no-hsts --no-check-certificate --content-disposition -N -P" & " " & Chr(34) & vLst(UBound(vLst)) & """", True, True, True, 0))
                           Call WriteActions("Downloading " & sFilename & " completed")
                           Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        
                        End If
                    
                    End If
                               
'---------------------------------------------------------------------------------
'Purpose : Kill Process (Used to terminate tasks by process id (PID) or image name)
'---------------------------------------------------------------------------------
                    Case "TaskKill"
                         
                        If IsEXERunning(sFilename) Then
                         Call WriteActions("Running process " & sFilename & " found. Please close before cleanup ...")
                    
                             If vLst(UBound(vLst)) = UCase("WARNING") Then 'Show optional Warning
                                 Respond = "Process " & UCase(sFilename) & " will be closed before running this script." & vbNewLine & _
                                 "Do you wish to continue?"
                                 
                                 answer = MsgBox(Respond, vbDefaultButton1 + vbYesNo, sFilename)
                    
                                    If answer = vbYes Then
                                    RetVal = ShellWait("cmd.exe /C" & "taskkill /f /im " & sFilename, False)
                                    Call WriteActions(frmMain.lstDS.Text)
                                    Call WriteActions("Process " & (sFilename) & " <closed>")
                                    Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                    
                                    Else
                             
                                    End If
                                    
                             Else
                             
                                'Kill process without hint
                                 RetVal = ShellWait("cmd.exe /C" & "taskkill /f /im " & sFilename, False)
                                 Call WriteActions("Process " & (sFilename) & " <closed>")
                                 
                             End If
                            
                        End If
                        
'---------------------------------------------------------------------------------
'Purpose : Show Msg from script file
'---------------------------------------------------------------------------------
                    Case "Msg"
                        MsgBox Replace(sFilename, "\n", vbCrLf), , frmMain.lstDS.Text
                        

                End Select
            End If
            
        Next fileop
    End If
    
    'Clear Up
    Set iVals = Nothing
    Erase vLst
    sVal = vbNullString
    sFilename = vbNullString
    sCmd = vbNullString

End Function

'---------------------------------------------------------------------------------
'Purpose : Run UNDO Operations
'---------------------------------------------------------------------------------

Private Function DoUndoOps(ByVal Selection As String)

Dim iVals As New Collection
Dim sCmd As String
Dim sFilename As String
Dim ClonePath As String
Dim vLst() As String
                
Set iVals = GetIniSelection(Selection)
    
    If (iVals.count > 0) Then
        For Each fileop In iVals
            vLst = Split(Trim(INIReadKey(Selection, fileop)), ",", , vbBinaryCompare)
            sCmd = Trim(vLst(0))
            
            If SafeUBound(vLst) >= 1 Then
            
                sFilename = ExpandPath(vLst(1))
                
                Select Case sCmd
                
'---------------------------------------------------------------------------------
'Purpose : Delete file
'---------------------------------------------------------------------------------
                    Case "rm"
                    
                        If vLst(UBound(vLst)) = UCase("RECURSE") Then
                                
                            Call WriteActions("[Delete file(s) /recursively] " & sFilename)
                            RetVal = ShellWait("cmd.exe /C" & "del /s /q " & Chr(34) & sFilename & Chr(34), False)
                            Call WriteActions("[Delete file /recursively] " & sFilename & " <done>")
                        Else
                             Call WriteActions("[Delete file] " & sFilename)
                             RetVal = ShellWait("cmd.exe /c" & "DEL /F /Q " & Chr(34) & sFilename & Chr(34), False)
                             Call WriteActions("[Delete file] " & sFilename & " <done>")
                        End If
                        
'---------------------------------------------------------------------------------
'Purpose : Delete folder
'---------------------------------------------------------------------------------
                Case "rmdir"
             

                    If vLst(UBound(vLst)) = UCase("RECURSE") Then
                        
                        Call WriteActions("[Delete directory /recursively] " & sFilename)
                        RetVal = ShellWait("cmd.exe /C" & "rmdir /s /q " & Chr(34) & sFilename & Chr(34), False)
                        Call WriteActions("[Delete directory /recursively] " & sFilename & " <done>")
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                           
                    Else
                        Call WriteActions("[Delete directory] " & sFilename)
                        RetVal = ShellWait("cmd.exe /C" & "DEL /F /Q " & Chr(34) & sFilename & Chr(34), False)
                         Call WriteActions("[Delete directory] " & sFilename & " <done>")
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
            
                    End If


'---------------------------------------------------------------------------------
'Purpose : Run file
'---------------------------------------------------------------------------------
                    Case "Start"
                    
                        Call WriteActions(frmMain.lstDS.Text)
                        Call WriteActions("[Run] " & sFilename & " <processing>")
                                              
                        Call ShellExecute(hwnd, "Open", _
                                 sFilename, "", "", 1)
                             
                        Call WriteActions("[Run] " & sFilename & " <done>")
                        Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                      
'---------------------------------------------------------------------------------
'Purpose : Start (executables)
'---------------------------------------------------------------------------------
                    Case "Start()"
                    
                        If DirExists(Environ("WinDir") & "\SysWOW64") Then  'x64 (Using the virtual "Sysnative" folder to access 64-bit cmd.exe from 32-bit code)
                            
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("[Start(Wait)] " & sFilename & " <processing>")
                                                  
                            RetVal = ShellWait(Environ("SystemRoot") & "\sysnative\cmd.exe /C" & Chr(34) & sFilename & Chr(34), False)
                                 
                            Call WriteActions("[Start(Wait)] " & sFilename & " <done>")
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    
                        Else 'x86-bit
                         
                            Call WriteActions(frmMain.lstDS.Text)
                            Call WriteActions("[Start(Wait)] " & sFilename & " <processing>")
    
                            RetVal = ShellWait("cmd.exe /C" & Chr(34) & sFilename & Chr(34), False)
                                   
                            Call WriteActions("[Start(Wait)] " & sFilename & " <done>")
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                            
                        End If
                                                      
                       
'---------------------------------------------------------------------------------
'Purpose : Download files with Wget
'---------------------------------------------------------------------------------
                  Case "Get"
                
                     If vLst(UBound(vLst)) = UCase("DEFAULT") Then 'Download to directory set in frmMain.txtOutputDir.Text
                                    
                       Call WriteActions(frmMain.lstDS.Text)
                       Call WriteActions(GetCommandOutput("" & Chr(34) & frmMain.txtWgetPath.Text & """" & " " & sFilename & " " & frmMain.txtWgetParam.Text & " " & Chr(34) & frmMain.txtOutputDir.Text & "", True, True, True, 0))
                       Call WriteActions("Downloading " & sFilename & " completed")
                       Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    
              
                   Else

                       Call WriteActions(frmMain.lstDS.Text)
                       Call WriteActions(GetCommandOutput("" & Chr(34) & frmMain.txtWgetPath.Text & """" & " " & sFilename & " " & frmMain.txtWgetParam.Text & " " & Chr(34) & vLst(UBound(vLst)) & "", True, True, True, 0))
                       Call WriteActions("Downloading " & sFilename & " completed")
                       Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    
                    End If
                    
'---------------------------------------------------------------------------------
'Purpose : Download files RAW with Wget paramter --content-disposition
'---------------------------------------------------------------------------------
                Case "GetRaw"
                
                    If vLst(UBound(vLst)) = UCase("DEFAULT") Then 'Download to directory set in frmMain.txtOutputDir.Text
                                    
                       Call WriteActions(frmMain.lstDS.Text)
                       Call WriteActions(GetCommandOutput("" & Chr(34) & frmMain.txtWgetPath.Text & """" & " " & Chr(34) & sFilename & "?raw=true" & """" & " " & "--show-progress --progress=bar:force --no-hsts --no-check-certificate --content-disposition -N -P" & " " & Chr(34) & frmMain.txtOutputDir.Text & """", True, True, True, 0))
                       Call WriteActions("Downloading " & sFilename & " completed")
                       Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    

                   Else

                       Call WriteActions(frmMain.lstDS.Text)
                       Call WriteActions(GetCommandOutput("" & Chr(34) & frmMain.txtWgetPath.Text & """" & " " & Chr(34) & sFilename & "?raw=true" & """" & " " & "--show-progress --progress=bar:force --no-hsts --no-check-certificate --content-disposition -N -P" & " " & Chr(34) & vLst(UBound(vLst)) & """", True, True, True, 0))
                       Call WriteActions("Downloading " & sFilename & " completed")
                       Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    
                    End If
                    
'---------------------------------------------------------------------------------
'Purpose : Run commands with reg.exe in command-line
'---------------------------------------------------------------------------------
                    Case "Reg"
                    
                    'x64 (Using the virtual "Sysnative" folder to access 64-bit reg.exe from 32-bit code)
                        'more Info here http://scottiestech.info/2009/06/20/calling-a-32-bit-system-command-from-a-script-in-x64-windows
                         If DirExists(Environ("WinDir") & "\SysWOW64") Then
                    
                            If vLst(UBound(vLst)) = UCase("STDOUT") Then 'Capture output
                            
                                   Call WriteActions("[Undo] " & frmMain.lstDS.Text)
                                   Call WriteActions("//Here the results of running")
                                   Call WriteActions("[Reg] " & sFilename)
                                   Call WriteActions(GetCommandOutput(Environ("SystemRoot") & "\sysnative\reg.exe" & " " & sFilename, True, True, True, 0))
                                   Call WriteActions("[Undo] " & frmMain.lstDS.Text & " <done>")
                                   Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                               
                            ElseIf vLst(UBound(vLst)) = UCase("SHOWCLI") Then ' Show optional Command prompt window
                            
                                   Call WriteActions("[Undo] " & frmMain.lstDS.Text)
                                   Call WriteActions("//Results are opened in command prompt")
                                   RetVal = ShellWait(Environ("SystemRoot") & "\sysnative\reg.exe /C" & " " & sFilename, True)
                                   Call WriteActions("[Undo] " & frmMain.lstDS.Text & " <done>")
                                   Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                               
                              Else  'Run silent
                              
                                   Call WriteActions("[Undo] " & frmMain.lstDS.Text)
                                   Call WriteActions("[Reg] " & sFilename & " <processing>")
                                   RetVal = ShellWait(Environ("SystemRoot") & "\sysnative\reg.exe /C" & " " & sFilename, False)
                                   Call WriteActions("[Undo] " & frmMain.lstDS.Text & " <done>")
                                   Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                
                                  
                              End If
                              
                        Else
                        
                         'x86
                    
                            If vLst(UBound(vLst)) = UCase("STDOUT") Then 'Capture output
                            
                                   Call WriteActions("[Undo] " & frmMain.lstDS.Text)
                                   Call WriteActions("//Here the results of running")
                                   Call WriteActions("[Reg] " & sFilename)
                                   Call WriteActions(GetCommandOutput("reg.exe" & " " & sFilename, True, True, True, 0))
                                   Call WriteActions("[Undo] " & frmMain.lstDS.Text & " <done>")
                                   Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                               
                            ElseIf vLst(UBound(vLst)) = UCase("SHOWCLI") Then ' Show optional Command prompt window
                            
                                   Call WriteActions("[Undo] " & frmMain.lstDS.Text)
                                   Call WriteActions("//Results are opened in command prompt")
                                   RetVal = ShellWait("cmd /K reg.exe" & " " & sFilename, True)
                                   Call WriteActions("[Undo] " & frmMain.lstDS.Text & " <done>")
                                   Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                               
                              Else  'Run silent
                              
                                   Call WriteActions("[Undo] " & frmMain.lstDS.Text)
                                   Call WriteActions("[Reg] " & sFilename & " <processing>")
                                   RetVal = ShellWait("reg.exe" & " " & sFilename, False)
                                   Call WriteActions("[Undo] " & frmMain.lstDS.Text & " <done>")
                                   Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                
                                  
                              End If
                    
                    End If
                    
'---------------------------------------------------------------------------------
'Purpose : Run commands in command-line
'---------------------------------------------------------------------------------
                    Case "Run"
        
                        If vLst(UBound(vLst)) = UCase("STDOUT") Then 'Capture output
                            
                            Call WriteActions("[Undo] " & frmMain.lstDS.Text)
                            Call WriteActions("//Here the results of running")
                            Call WriteActions("[Run] " & sFilename & " <processing>")
                            Call WriteActions(GetCommandOutput(sFilename, True, True, True, 0))
                            Call WriteActions("[Undo] " & frmMain.lstDS.Text & " <done>")
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                      
                        ElseIf vLst(UBound(vLst)) = UCase("SHOWCLI") Then ' Show optional Command prompt window
                                   
                            Call WriteActions("[Undo] " & frmMain.lstDS.Text)
                            Call WriteActions("//Results are opened in command prompt")
                            RetVal = ShellWait("cmd.exe /C" & sFilename, True)
                            Call WriteActions("[Undo] " & frmMain.lstDS.Text & " <done>")
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                      
                        Else  'Run silent
                                     
                            Call WriteActions("[Undo] " & frmMain.lstDS.Text)
                            Call WriteActions("[Run] " & sFilename & " <processing>")
                            RetVal = ShellWait("cmd.exe /C" & sFilename, False)
                            Call WriteActions("[Undo] " & frmMain.lstDS.Text & " <done>")
                            Call WriteActions("; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                               
                        End If
                        
'---------------------------------------------------------------------------------
'Purpose : Show Msg from script file (E.g., If UNDO Operations are not possible)
'---------------------------------------------------------------------------------
                    Case "Msg"
                        MsgBox Replace(sFilename, "\n", vbCrLf), , frmMain.lstDS.Text

                End Select
            End If
            
        Next fileop
    End If
    
    'Clear Up
    Set iVals = Nothing
    Erase vLst
    sVal = vbNullString
    sFilename = vbNullString
    sCmd = vbNullString

End Function
Public Property Get IsLoaded() As Boolean
    IsLoaded = FindFile(m_IniFilename)
End Property

Public Property Get scID() As String
    If (IsLoaded) Then
        scID = INIReadKey("Info", "ID")
    End If
End Property
Public Property Get scDesc() As String
    If (IsLoaded) Then
        scDesc = Replace(INIReadKey("Info", "Desc"), "\n", vbCrLf)
    End If
End Property
Public Property Get scVer() As String
    If (IsLoaded) Then
        scVer = INIReadKey("Info", "Ver")
    End If
End Property
Public Property Get scDev() As String
    If (IsLoaded) Then
        scDev = INIReadKey("Info", "Dev")
    End If
End Property
Public Property Get scDevURL() As String
    If (IsLoaded) Then
        scDevURL = INIReadKey("Info", "DevURL")
    End If
End Property
Public Property Get scWarning() As String
    If (IsLoaded) Then
        scWarning = INIReadKey("Info", "Warning")
    End If
End Property
Public Property Get scEvaluation() As String
    If (IsLoaded) Then
        scEvaluation = INIReadKey("Info", "Evaluation")
    End If
End Property
Public Property Get scEvaluationColor() As String
    If (IsLoaded) Then
        scEvaluationColor = INIReadKey("Info", "EvaluationColor")
    End If
End Property
Public Property Get scUpdate() As String
    If (IsLoaded) Then
        scUpdate = INIReadKey("Info", "Update")
    End If
End Property
Public Property Get scRequireAdmin() As String
    If (IsLoaded) Then
        scRequireAdmin = INIReadKey("Info", "RequireAdmin")
    End If
End Property
Public Property Get scPatron() As String
    If (IsLoaded) Then
        scPatron = INIReadKey("Info", "Patron")
    End If
End Property
