Attribute VB_Name = "modFolderBrowser"
Option Explicit

'---------------------------------------------------------------------------------
' Copyright!
' User control:  modUniBrowseForFolder
' Author:        Zhu JinYong
' Original By:   DaVBMan (http://vbcity.com/forums/t/67223.aspx)
' Dependencies:  -None-
' Compatibility: Unicode Tested on XP/Vista/Win7/Win8/Win10
' Last revision: 18/5/2011
'---------------------------------------------------------------------------------

Private Type BrowseInfo
   hwndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long 'String
   lpszTitle As Long 'String
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_EDITBOX = &H10
Public Const BIF_VALIDATE = &H20
Public Const BIF_NEWDIALOGSTYLE = &H40
Public Const BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
Public Const BIF_BROWSEINCLUDEURLS = &H80
Public Const BIF_UAHINT = &H100
Public Const BIF_NONEWFOLDERBUTTON = &H200
Public Const BIF_NOTRANSLATETARGETS = &H400
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
Public Const BIF_BROWSEINCLUDEFILES = &H4000
Public Const BIF_SHAREABLE = &H8000
Private Const MAX_PATH = 260
Private Const WM_USER = &H400

Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_VALIDATEFAILEDA = 3  'lParam:szPath ret:1(cont),0(EndDialog)
Private Const BFFM_VALIDATEFAILEDW = 4  'lParam:wzPath ret:1(cont),0(EndDialog)
Private Const BFFM_IUNKNOWN = 5         'provides IUnknown to client. lParam: IUnknown*

'messages to browser
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSTATUSTEXTW = WM_USER + 104
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW = (WM_USER + 103)
Private Const BFFM_ENABLEOK = WM_USER + 101

Public Declare Function SHGetPathFromIDListA Lib "shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetPathFromIDListW Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As Long) As Long

Public Declare Function SHBrowseForFolderA Lib "shell32.dll" (lpBrowseInfo As BrowseInfo) As Long
Private Declare Function SHBrowseForFolderW Lib "shell32" (lpbi As BrowseInfo) As Long

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    
Public Function BrowseForFolder(ByVal hWndModal As Long, Optional StartFolder As String = "", Optional Title As String = "Please select a folder:", _
   Optional IncludeFiles As Boolean = False, Optional IncludeNewFolderButton As Boolean = False) As String
    Dim bInf As BrowseInfo
    Dim RetVal As Long
    Dim PathID As Long
    Dim RetPath As String
    Dim Offset As Integer
    Dim szTitleInfo() As Byte
    Dim strSTARTFOLDER As String
    
    'Set the properties of the folder dialog
    With bInf
         .hwndOwner = hWndModal
         .pIDLRoot = 0
         szTitleInfo = Title & vbNullChar
         .lpszTitle = VarPtr(szTitleInfo(0))
         .ulFlags = IIf(IncludeFiles, BIF_BROWSEINCLUDEFILES, BIF_RETURNONLYFSDIRS) + BIF_DONTGOBELOWDOMAIN + BIF_USENEWUI + _
                    IIf(IncludeNewFolderButton, 0&, BIF_NONEWFOLDERBUTTON)
         If IncludeFiles Then .ulFlags = .ulFlags Or BIF_BROWSEINCLUDEFILES
         If IncludeNewFolderButton Then .ulFlags = .ulFlags Or BIF_NEWDIALOGSTYLE
         If StartFolder <> "" Then
            strSTARTFOLDER = StartFolder & vbNullChar
           .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc) 'get address of function.
           .lParam = StrPtr(strSTARTFOLDER)
         End If
     End With
     
    'Show the Browse For Folder dialog
    PathID = SHBrowseForFolderW(bInf)
    If PathID = 0 Then Exit Function
    RetPath = Space$(MAX_PATH)
    RetVal = SHGetPathFromIDListW(PathID, StrPtr(RetPath))
    If RetVal Then
         'Trim off the null chars ending the path
         'and display the returned folder
         Offset = InStr(RetPath, Chr$(0))
         BrowseForFolder = Left$(RetPath, Offset - 1)
         'Free memory allocated for PIDL
         CoTaskMemFree PathID
    Else
         BrowseForFolder = ""
    End If
End Function
Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
   On Error Resume Next
   Dim lpIDList As Long
   Dim ret As Long
   Dim sBuffer As String
   Select Case uMsg
       Case BFFM_INITIALIZED
           Call SendMessageW(hwnd, BFFM_SETSELECTIONW, 1&, pData) 'StrPtr(mstrSTARTFOLDER)) 'Private mstrSTARTFOLDER As String
       Case BFFM_SELCHANGED
           sBuffer = Space(MAX_PATH)
           ret = SHGetPathFromIDListW(lp, StrPtr(sBuffer))
           If ret = 1 Then
               Call SendMessageW(hwnd, BFFM_SETSTATUSTEXTW, 0, StrPtr(sBuffer))
           End If
   End Select
   BrowseCallbackProc = 0
End Function

Private Function GetAddressofFunction(add As Long) As Long
   GetAddressofFunction = add
End Function



