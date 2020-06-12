Attribute VB_Name = "modEditor"
'---------------------------------------------------------------------------------
'Public  Declaration  : Code editor
'---------------------------------------------------------------------------------

Private Declare Function GetTempFileName Lib "kernel32" _
  Alias "GetTempFileNameA" ( _
  ByVal lpszPath As String, _
  ByVal lpPrefixString As String, _
  ByVal wUnique As Long, _
  ByVal lpTempFileName As String) As Long
 
Private Declare Function GetTempPath Lib "kernel32.dll" _
  Alias "GetTempPathA" ( _
  ByVal nBufferLength As Long, _
  ByVal lpBuffer As String) As Long
 
Private winTempPath As String
 
'---------------------------------------------------------------------------------
'Purpose  : Determines a temporary file name in the temporary directory
'---------------------------------------------------------------------------------

Private Function txt_TempFilename() As String
  Dim myTempFileName As String
  Dim RetVal As Long
 
  If winTempPath = "" Then
    ' Determines a temporary directory
    winTempPath = Space$(256)
    RetVal = GetTempPath(Len(winTempPath), winTempPath)
    winTempPath = Left$(winTempPath, RetVal)
  End If
 
  ' Determines a temporary file
  myTempFileName = Space$(256)
  Call GetTempFileName(winTempPath, "txt", 0&, myTempFileName)
  myTempFileName = Left$(myTempFileName, _
    InStr(myTempFileName, Chr$(0)) - 1)
 
  txt_TempFilename = myTempFileName
End Function

'---------------------------------------------------------------------------------
'Purpose  : Read the entire contents of a text file
'---------------------------------------------------------------------------------

Public Function txt_ReadAll(ByVal sFilename As String) _
  As String
 
  Dim F As Integer
  Dim sInhalt As String
 
  ' File exists?
  If Dir$(sFilename, vbNormal) <> "" Then
    ' Open text file in binary mode and total and read content in one go
    
    F = FreeFile
    Open sFilename For Binary As #F
    sInhalt = Space$(LOF(F))
    Get #F, , sInhalt
    Close #F
  End If
 
  txt_ReadAll = sInhalt
  
End Function

'---------------------------------------------------------------------------------
' Purpose   :  Save any text in a text file
'              The previous contents of the text file is entirety overwritten!
'---------------------------------------------------------------------------------

Public Sub txt_WriteAll(ByVal sFilename As String, _
  ByVal sLines As String)
 
  Dim F As Integer
 
   ' Open file for writing
   ' Warning: previous content will be deleted!
   
  F = FreeFile
  Open sFilename For Output As #F
  Print #F, sLines
  Close #F
  
End Sub
