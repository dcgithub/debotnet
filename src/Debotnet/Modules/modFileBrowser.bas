Attribute VB_Name = "modFileBrowser"
'---------------------------------------------------------------------------------
'Public  Declaration  : This Function sets the Filters for the Common Dialog.
'                       It is basically the same as in CommonDialog OCX
'---------------------------------------------------------------------------------

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private strfileName As OPENFILENAME

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    
End Type

Private Sub DialogFilter(WantedFilter As String)

Dim intLoopCount As Integer
strfileName.lpstrFilter = ""
For intLoopCount = 1 To Len(WantedFilter)
If Mid(WantedFilter, intLoopCount, 1) = "|" Then strfileName.lpstrFilter = _
strfileName.lpstrFilter + Chr(0) Else strfileName.lpstrFilter = _
strfileName.lpstrFilter + Mid(WantedFilter, intLoopCount, 1)
Next intLoopCount
strfileName.lpstrFilter = strfileName.lpstrFilter + Chr(0)

End Sub

'---------------------------------------------------------------------------------
'Function to get the Filename to Open
'---------------------------------------------------------------------------------
Public Function fncGetFileNametoOpen(Optional strDialogTitle As String = "Open", Optional strFilter As String = "All Files|*.*", Optional strDefaultExtention As String = "*.*") As String

Dim lngReturnValue As Long
Dim intRest As Integer
strfileName.lpstrTitle = strDialogTitle
strfileName.lpstrDefExt = strDefaultExtention
DialogFilter (strFilter)
strfileName.hInstance = App.hInstance
strfileName.lpstrFile = Chr(0) & Space(259)
strfileName.nMaxFile = 260
strfileName.flags = &H4
strfileName.lStructSize = Len(strfileName)
lngReturnValue = GetOpenFileName(strfileName)
fncGetFileNametoOpen = strfileName.lpstrFile

End Function

'---------------------------------------------------------------------------------
'This Function returns the Save File Name
'---------------------------------------------------------------------------------
Public Function fncGetFileNametoSave(strFilter As String, strDefaultExtention As String, Optional strDialogTitle As String = "Save") As String

Dim lngReturnValue As Long
Dim intRest As Integer
strfileName.lpstrTitle = strDialogTitle
strfileName.lpstrDefExt = strDefaultExtention
DialogFilter (strFilter)
strfileName.hInstance = App.hInstance
strfileName.lpstrFile = Chr(0) & Space(259)
strfileName.nMaxFile = 260
strfileName.flags = &H80000 Or &H4
strfileName.lStructSize = Len(strfileName)
lngReturnValue = GetSaveFileName(strfileName)
fncGetFileNametoSave = strfileName.lpstrFile

End Function

