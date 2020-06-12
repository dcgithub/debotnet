Attribute VB_Name = "modGlobal"
'---------------------------------------------------------------------------------
'Public  Declaration  : Detect right-click in frmmain.lstDS
'---------------------------------------------------------------------------------
 
Public Const LB_GETTOPINDEX = &H18E
Public Const LB_GETITEMHEIGHT = &H1A1

'---------------------------------------------------------------------------------
'Purpose : Autocomplete in Search > frmmain.txtSearch
'---------------------------------------------------------------------------------

Public Const LB_FINDSTRING = &H18F
Public Declare Function SendMessage Lib _
                    "user32" Alias "SendMessageA" (ByVal _
                    hwnd As Long, ByVal wMsg As Long, _
                    ByVal wParam As Long, lParam As Any) _
                    As Long
                    
'---------------------------------------------------------------------------------
' Public Declaration  :  Starts an application or a document with the linked application
'---------------------------------------------------------------------------------

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'---------------------------------------------------------------------------------
' Public Declaration  :  Check if file exists (Unicode ver)
'---------------------------------------------------------------------------------

Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long

'---------------------------------------------------------------------------------
' Purpose :  Wait until app has been closed
'---------------------------------------------------------------------------------

Private Declare Function CreateProcess Lib "kernel32" Alias _
                                            "CreateProcessA" ( _
    ByVal lpAppName As Long, _
    ByVal lpCmdLine As String, _
    ByVal lpProcAttr As Long, _
    ByVal lpThreadAttr As Long, _
    ByVal lpInheritedHandle As Long, _
    ByVal lpCreationFlags As Long, _
    ByVal lpEnv As Long, _
    ByVal lpCurDir As Long, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInfo As PROCESS_INFORMATION _
    ) As Long
     
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    
Private Const NORMAL_PRIORITY_CLASS  As Long = &H20&
Private Const INFINITE As Long = -1&
Private Const WAIT_TIMEOUT As Long = 258&

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Integer
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

'---------------------------------------------------------------------------------
' Purpose :  Show Open with dialog
'---------------------------------------------------------------------------------

Public Declare Function GetSystemDirectory Lib "kernel32" _
  Alias "GetSystemDirectoryA" ( _
  ByVal lpBuffer As String, _
  ByVal nSize As Long) As Long
 
Public Declare Function GetDesktopWindow Lib "user32" () As Long
 
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_NOTFOUND = 2

'---------------------------------------------------------------------------------
'Purpose : Show Open file with dialog
'---------------------------------------------------------------------------------

Public Sub DocumentOpen(sFilename As String)
  Dim sDirectory As String
  Dim lRet As Long
  Dim DeskWin As Long
 
  DeskWin = GetDesktopWindow()
  lRet = ShellExecute(DeskWin, "open", sFilename, _
    vbNullString, vbNullString, vbNormalFocus)
 
  If lRet = SE_ERR_NOTFOUND Then
    ' file not found
 
  ElseIf lRet = SE_ERR_NOASSOC Then
    ' If filetype is unknow, show "Open with" dialog
    
    sDirectory = Space(260)
    lRet = GetSystemDirectory(sDirectory, Len(sDirectory))
    sDirectory = Left(sDirectory, lRet)
    Call ShellExecute(DeskWin, vbNullString, _
      "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & _
      sFilename, sDirectory, vbNormalFocus)
  End If
  
End Sub

'---------------------------------------------------------------------------------
' Purpose :  Wait until app has been closed
'---------------------------------------------------------------------------------

Public Function ShellWait(cmdline As String, Optional ByVal _
                        bShowApp As Boolean = False) As Boolean

    'Reserve memory
    Dim uProc As PROCESS_INFORMATION
    Dim uStart As STARTUPINFO
    Dim lRetVal As Long
    
    'Initialize Data types
    uStart.cb = Len(uStart)
    uStart.wShowWindow = Abs(bShowApp)
    uStart.dwFlags = 1
    
    'Create Process
    lRetVal = CreateProcess(0&, cmdline, 0&, 0&, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, uStart, uProc)
    
    If lRetVal = 0 Then
        Call WriteActions("Failed starting executable in script " & frmMain.lstDS.Text)
        
        ShellWait = False
        Exit Function
    End If
    
    
    'Wait until Process has been closed
    'Refresh Process
    Do While WaitForSingleObject(uProc.hProcess, 10) = WAIT_TIMEOUT
        DoEvents
    Loop
    
    'Force to wait and not only stillstand
    'lRetVal = WaitForSingleObject(uProc.hProcess, INFINITE)
    
    'Close Process
    lRetVal = CloseHandle(uProc.hProcess)

    ShellWait = (lRetVal <> 0)
End Function

'---------------------------------------------------------------------------------
' Purpose :  Check if file exists method
'---------------------------------------------------------------------------------

Public Function FileExists(ByRef sFilename As String) As Boolean

    Const ERROR_SHARING_VIOLATION = 32&

    Select Case (GetFileAttributesW(StrPtr(sFilename)) And vbDirectory) = 0&
        Case True: FileExists = True
        Case Else: FileExists = Err.LastDllError = ERROR_SHARING_VIOLATION
        
    End Select
    
End Function

'---------------------------------------------------------------------------------
' Purpose :  Check if directory exists (with Wildcard support)
'---------------------------------------------------------------------------------

Public Function DirExists(ByVal sPath As String) As Boolean

On Error Resume Next

  Dim iResult As Integer
  Dim sDir As String
 

  sDir = Dir$(sPath, vbDirectory Or vbHidden Or vbSystem)
  If Len(sDir) > 0 Then
    If InStr(sPath, "\") > 0 Then
      sPath = Left$(sPath, InStrRev(sPath, "\"))
      sDir = sPath & sDir
    End If
    iResult = GetAttr(sDir)
    DirExists = IIf(Err = 0, True, False)
  End If
  
End Function

'---------------------------------------------------------------------------------
' Purpose :  Check if directory exists
' PIPELINE!
'---------------------------------------------------------------------------------

Public Function DirExistsEx(ByVal DirectoryName As String) As Boolean
    On Error Resume Next
    DirExistsEx = CBool(GetAttr(DirectoryName) And vbDirectory)
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------
' Purpose :  Used to find the last backslash of the file path
'---------------------------------------------------------------------------------

Public Function GetLastSlash(Text As String) As String

On Error GoTo Err

Dim i, Pos As Integer
Dim LastSlash As Integer

For i = 1 To Len(Text)
Pos = InStr(i, Text, "/", vbTextCompare)
If Pos <> 0 Then LastSlash = Pos
Next i
GetLastSlash = Right(Text, Len(Text) - LastSlash)

Err:
Exit Function
End Function

