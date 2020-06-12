Attribute VB_Name = "modUpdate"

'---------------------------------------------------------------------------------
'Purpose  : Check for updates (Stable and nightly releases)
'---------------------------------------------------------------------------------

Public Function CheckAppUpdate()

On Error Resume Next 'If version.txt not found on server

    Dim VersionDebotnet As String
    Dim VersionDebotnetINI As String
    
    Dim sURL As String
    Dim sLocalFile As String
    
'---------------------------------------------------------------------------------
'Add here the path to latest version.txt file on your server!!!
    sURL = "http://path-to-version.txt"
'---------------------------------------------------------------------------------

    ' Filename on local system
    sLocalFile = App.Path & "\bin\"
    ' Download file
    RetVal = ShellWait("cmd.exe /C" & """" & Chr(34) & frmMain.txtWgetPath.Text & """" & " " & Chr(34) & sURL & """" & " " & "-q --show-progress --progress=bar:force --no-hsts --no-check-certificate --content-disposition -N -P" & " " & Chr(34) & sLocalFile & """", False)

'---------------------------------------------------------------------------------
'Release channel STABLE
'---------------------------------------------------------------------------------

If frmMain.mnuReleaseStable.Checked = True Then

    'Version Debotnet
    VersionDebotnet = App.Major & "." & App.Minor & "." & App.Revision
    VersionDebotnetINI = GetINIString(App.Path & "\bin\" & "version.txt", "Debotnet", "Version")
     
    If VersionDebotnetINI > VersionDebotnet Then
         
    'Update available!
    
          Msg = "Do you want to get version " & VersionDebotnetINI & " of Debotnet?" & vbNewLine & _
          "Click " & Chr(34) & "Cancel" & Chr(34) & " to read the release notes only."
              
            Response = MsgBox(Msg, vbYesNoCancel + vbDefaultButton1, "Update available")
            
            Select Case Response
            
                Case vbYes
         
                  frmMain.txtStatus.Text = ""
                  frmMain.txtStatus.Visible = True
       
                  frmMain.Caption = frmMain.Caption & " - downloading " & "(" & VersionDebotnetINI & ")" & " update"
                  
                  frmMain.txtStatus.Text = GetCommandOutput("" & Chr(34) & frmMain.txtWgetPath.Text & """" & " " & _
                  "https://github.com/mirinsoft/debotnet/releases/download/" & VersionDebotnetINI & "/debotnet.zip" & " " & " " & "-q --show-progress --progress=bar:force --no-hsts --no-check-certificate -N -O" & " " & Chr(34) & "debotnet_" & VersionDebotnetINI & ".zip" & "", True, True, True, 0)
                  
                  frmMain.Caption = "Debotnet " & "- " & "downloading update - completed."
                      
                  frmMain.txtStatus.Visible = False
                  
                  MsgBox "Update saved to " & App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "debotnet_" & VersionDebotnetINI & ".zip"
                    
                    
                 Case vbNo
                 
                 
                 Case vbCancel
                 
                     'Show release notes on GitHub
                        Call ShellExecute(hwnd, "Open", _
                        "https://github.com/mirinsoft/debotnet/releases/tag/" & VersionDebotnetINI, "", "", 1)
    
             End Select
             
        
        Kill App.Path & "\bin\version.txt" 'delete version.txt again
    
              
    ElseIf VersionDebotnetINI = VersionDebotnet Then
        
    
        MsgBox "You are using the latest version of Debotnet (Pegasos).", vbInformation, App.EXEName & " (Stable)"
        
        Kill App.Path & "\bin\version.txt" 'delete version.txt again
         
    ElseIf VersionDebotnetINI < VersionDebotnet Then
      
    
        MsgBox "You are using an unofficial version of Debotnet.", vbDefaultButton1, App.EXEName & " (Stable)"
        
        Kill App.Path & "\bin\version.txt" 'delete version.txt again
             
    End If
    
Else

'---------------------------------------------------------------------------------
'Release channel NIGHLTY
'---------------------------------------------------------------------------------

    'Version Debotnet
    VersionDebotnet = App.Major & "." & App.Minor & "." & App.Revision
    VersionDebotnetINI = GetINIString(App.Path & "\bin\" & "version.txt", "Debotnet", "Nightly")
     
    If VersionDebotnetINI > VersionDebotnet Then
         
    'Update available!
    
          Msg = "Do you want to get version " & VersionDebotnetINI & " of Debotnet?"
              
            Response = MsgBox(Msg, vbYesNo + vbDefaultButton1, "Nightly build available")
            
            Select Case Response
            
                Case vbYes
         
                   Call ShellExecute(hwnd, "Open", _
                         "https://www.mirinsoft.com/debotnet-nightly", "", "", 1)
      
                 Case vbNo
                 
    
             End Select
             
        
        Kill App.Path & "\bin\version.txt" 'delete version.txt again
    
              
    Else
        
    
        MsgBox "There are currently no updates available.", vbInformation, App.EXEName & " (Nightly)"
        
        Kill App.Path & "\bin\version.txt" 'delete version.txt again
         
             
    End If
    
End If

End Function

