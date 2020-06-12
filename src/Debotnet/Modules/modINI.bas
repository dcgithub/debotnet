Attribute VB_Name = "modINI"
'---------------------------------------------------------------------------------
'Public  Declaration  : Read/Write INI (Config file Debotnet)
'---------------------------------------------------------------------------------

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal sSectionName As String, ByVal sKeyName As String, ByVal sDefault As String, ByVal sReturnedString As String, ByVal lSize As Long, ByVal sFilename As String) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'---------------------------------------------------------------------------------
'Purpose : Write INI (Config file Debotnet > Custom Settings)
'---------------------------------------------------------------------------------

Public Function WriteINI(FileName As String, Section As String, _
  Name As String, Value As String) As Long
   WriteINI = WritePrivateProfileString(Section, Name, _
   Value, FileName)
End Function

'---------------------------------------------------------------------------------
'Purpose : Read INI (Config file Debotnet > Custom Settings)
'---------------------------------------------------------------------------------

Public Function GetINIString(FileName As String, _
   Section As String, Name As String) As String
   temp$ = String(255, 0)
   X = GetPrivateProfileString(Section, Name, "", _
             temp$, 255, FileName)
   temp$ = Left$(temp$, X)
   GetINIString = temp$
End Function

'---------------------------------------------------------------------------------
'Purpose : Read INI Values (Config file debotnet-settings.txt)
'---------------------------------------------------------------------------------

Private Function ReadValue(ByVal strSectionName As String, ByVal strKeyName As String, ByVal strDefaultValue As String, ByVal strfileName As String) As String

Dim lngBufferSize As Long
Dim lngLength As Long
Dim strBuffer As String
    strBuffer = String$(2000, 0)
    lngBufferSize = Len(strBuffer)
    lngLength = GetPrivateProfileString(strSectionName, strKeyName, strDefaultValue, strBuffer, lngBufferSize, strfileName)
    strBuffer = Left$(strBuffer, lngLength)
    ReadValue = strBuffer
    
End Function

'---------------------------------------------------------------------------------
'Purpose : Write INI Values (Config file debotnet-settings.txt)
'---------------------------------------------------------------------------------

Public Sub WriteValue(ByVal strSectionName As String, ByVal strKeyName As String, ByVal strValue As String, ByVal strfileName As String)

    WritePrivateProfileString strSectionName, strKeyName, strValue, strfileName
    
End Sub

'---------------------------------------------------------------------------------
'Purpose : Load individual configuration settings from debotnet-settings.txt
'---------------------------------------------------------------------------------

Public Function LoadSettings()
        
On Error Resume Next

'lstDS
FileName$ = App.Path
If Right(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
FileName$ = FileName$ & "\bin\debotnet-settings.txt"
    
'//Main
   frmMain.lstCS.ListIndex = CStr(GetINIString(FileName$, "Settings", "Category")) 'Default category
   frmMain.lstTheme.Text = CStr(GetINIString(FileName$, "Settings", "Design")) 'Default theme

   frmMain.mnuReleaseStable.Checked = CStr(GetINIString(FileName$, "Release", "Stable")) 'Release channel Stable
   frmMain.mnuReleaseNightly.Checked = CStr(GetINIString(FileName$, "Release", "Nightly")) 'Release channel Nightly
   
'//Settings
   frmMain.txtRepository.Text = CStr(GetINIString(FileName$, "Repository", "URL")) 'Default repository
     
   frmMain.txtOutputDir.Text = CStr(GetINIString(FileName$, "Wget", "OutputDir")) 'Download directory
   frmMain.txtWgetPath.Text = CStr(GetINIString(FileName$, "Wget", "WgetPath")) 'Path to Wget.exe
   frmMain.txtWgetParam.Text = CStr(GetINIString(FileName$, "Wget", "WgetParam")) 'Wget parameter
   
   'Internal Editor Integration
   frmMain.chkUseDebotnetEditor.Value = CStr(GetINIString(FileName$, "Settings", "UseDebotnetEditor"))
   
   'Run always as Administrator
   frmMain.chkRunAlwaysInElevatedMode = CStr(GetINIString(FileName$, "Settings", "RunAlwaysInElevatedMode"))
   
    
End Function

'---------------------------------------------------------------------------------
'Purpose : Save individual configuration settings to debotnet-settings.txt
'---------------------------------------------------------------------------------

Public Function SaveSettings()

'debotnet-settings.txt
FileName$ = App.Path
If Right(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
FileName$ = FileName$ & "\bin\debotnet-settings.txt"
         
'//Main
    X = WriteINI(FileName$, "Settings", "Category", frmMain.lstCS.ListIndex) 'Default category
    X = WriteINI(FileName$, "Settings", "Design", frmMain.lstTheme.ListIndex) 'Default theme
    
    X = WriteINI(FileName$, "Release", "Stable", frmMain.mnuReleaseStable.Checked) 'Release channel Stable
    X = WriteINI(FileName$, "Release", "Nightly", frmMain.mnuReleaseNightly.Checked) 'Release channel Nightly
    
'//Settings
    X = WriteINI(FileName$, "Repository", "URL", frmMain.txtRepository.Text) 'Default repository
    
    X = WriteINI(FileName$, "Wget", "OutputDir", frmMain.txtOutputDir.Text) 'Download directory
    X = WriteINI(FileName$, "Wget", "WgetPath", frmMain.txtWgetPath.Text) 'Path to Wget.exe
    X = WriteINI(FileName$, "Wget", "WgetParam", frmMain.txtWgetParam.Text) 'Wget parameter
     
    'Internal Editor Integration
    X = WriteINI(FileName$, "Settings", "UseDebotnetEditor", frmMain.chkUseDebotnetEditor.Value)
    
   'Run always as Administrator
    X = WriteINI(FileName$, "Settings", "RunAlwaysInElevatedMode", frmMain.chkRunAlwaysInElevatedMode.Value)
    
End Function

'---------------------------------------------------------------------------------
'Purpose : Load Scripts from .INI File
'---------------------------------------------------------------------------------

Public Function LoadDSSettings()

IniFile$ = App.Path & "\bin\debotnet-settings.txt"

    Dim strSelection() As String
    strSelection = Split(ReadValue("Scripts", frmMain.lstCS, vbNullString, IniFile), ";")
    If UBound(strSelection) > 0 Then
        For j = 0 To UBound(strSelection)
            frmMain.lstDS.Selected(j) = CBool(strSelection(j))
        Next
    End If
            
End Function
 
'---------------------------------------------------------------------------------
'Purpose : Save Scripts to .INI File
'---------------------------------------------------------------------------------

Public Function SaveDSSettings()
        
    Dim j As Long
    Dim strSelection As String
    
    IniFile$ = App.Path & "\bin\debotnet-settings.txt"
         
    For j = 0 To frmMain.lstDS.ListCount - 1
        strSelection = strSelection & ";" & CInt(frmMain.lstDS.Selected(j))
    Next
    WriteValue "Scripts", frmMain.lstCS, Mid$(strSelection, 2), IniFile
     
End Function

'---------------------------------------------------------------------------------
'Purpose : Save window state and positions
'---------------------------------------------------------------------------------

Public Function SaveWindowState()

    If WritePrivateProfileString("WindowState", "Left", Format$(frmMain.Left), App.Path & "\" & "bin\debotnet-settings.txt") <= 0 Then
    End If
    If WritePrivateProfileString("WindowState", "Top", Format$(frmMain.Top), App.Path & "\" & "bin\debotnet-settings.txt") <= 0 Then
    End If
    If WritePrivateProfileString("WindowState", "Height", Format$(frmMain.Height), App.Path & "\" & "bin\debotnet-settings.txt") <= 0 Then
    End If
    If WritePrivateProfileString("WindowState", "Width", Format$(frmMain.Width), App.Path & "\" & "bin\debotnet-settings.txt") <= 0 Then
    End If
    
End Function

'---------------------------------------------------------------------------------
'Purpose : Load window state and positions
'---------------------------------------------------------------------------------
  
Public Function LoadWindowState()

On Error Resume Next
  
 Dim sRet As String
    Dim lRet As Long

    sRet = String$(60, 0)
    lRet = GetPrivateProfileString("WindowState", "Left", frmMain.StartUpPosition = 0, sRet, 60, App.Path & "\" & "bin\debotnet-settings.txt")
    If lRet > 0 Then
        sRet = Left$(sRet, lRet)
        If IsNumeric(sRet) Then
            frmMain.Left = CInt(sRet)
        End If
        Else
    End If
    
    sRet = String$(60, 0)
    lRet = GetPrivateProfileString("WindowState", "Top", frmMain.StartUpPosition = 0, sRet, 60, App.Path & "\" & "bin\debotnet-settings.txt")
    If lRet > 0 Then
        sRet = Left$(sRet, lRet)
        If IsNumeric(sRet) Then
             frmMain.Top = CInt(sRet)
        End If
        Else
    End If
    
        sRet = String$(60, 0)
    lRet = GetPrivateProfileString("WindowState", "Height", frmMain.StartUpPosition = 0, sRet, 60, App.Path & "\" & "bin\debotnet-settings.txt")
    If lRet > 0 Then
        sRet = Left$(sRet, lRet)
        If IsNumeric(sRet) Then
            frmMain.Height = CInt(sRet)
        End If
        Else
    End If
    
        sRet = String$(60, 0)
    lRet = GetPrivateProfileString("WindowState", "Width", frmMain.StartUpPosition = 0, sRet, 60, App.Path & "\" & "bin\debotnet-settings.txt")
    If lRet > 0 Then
        sRet = Left$(sRet, lRet)
        If IsNumeric(sRet) Then
             frmMain.Width = CInt(sRet)
        End If
        Else
    End If
    
End Function
