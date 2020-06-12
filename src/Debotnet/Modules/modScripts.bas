Attribute VB_Name = "modScripts"
'---------------------------------------------------------------------------------
'Public  Declaration  : Function to load scripts
'---------------------------------------------------------------------------------

Public m_ScriptCls As New clsScript
Public m_ScriptPath As String
Public m_ScriptCol As Collection

'---------------------------------------------------------------------------------
'Purpose : Function to load prefined categories
'---------------------------------------------------------------------------------

Public Function LoadCS()

On Error Resume Next

Dim Pfad As String, Name As String
Dim X As Long
    
    Attr = vbNormal Or vbReadOnly Or vbDirectory
  
    frmMain.lstCS.Visible = True
    frmMain.lstCS.Clear
    
    Pfad = App.Path & "\scripts\"
    Name = Dir$(Pfad, Attr)

    
    Do While Name <> ""
        If Name <> "." And Name <> ".." Then
            X = GetAttr(Pfad & Name)
            
            If (X And vbDirectory) <> 0 Then
            
               frmMain.lstCS.AddItem LCase(Replace(Name, "#", ""))
               
            End If
            
      End If
      
      Name = Dir(, Attr)
    Loop
    
End Function

'---------------------------------------------------------------------------------
'Purpose : Function to Load Scripts
'---------------------------------------------------------------------------------

Public Function LoadDS()

    If InStr(frmMain.lstCS.Text, "start") > 0 Then
    m_ScriptPath = FixPath(App.Path) & "scripts\" & "#" & frmMain.lstCS.Text & "\"
    Else
        m_ScriptPath = FixPath(App.Path) & "scripts\" & frmMain.lstCS.Text & "\"
    End If

    'Add scrip files to frmmain.lstDS
    Call LoadDebotnet
    
End Function

'---------------------------------------------------------------------------------
'Purpose : Load Debotnet script files
'---------------------------------------------------------------------------------

Private Sub LoadDebotnet()

On Error Resume Next

Dim xFile As String
Dim Pos1 As Long

'Create new collection object
Set m_ScriptCol = New Collection
        
    xFile = Dir(m_ScriptPath & "*.ds1", vbDirectory)
    
    Do Until (xFile = "")
    
    frmMain.lstDS.Visible = False 'Hide list for faster loading
    
    'Add Scripts to Collection
    m_ScriptCol.add m_ScriptPath & xFile
        
        'Add Scripts without Filetype
        Pos1 = InStrRev(xFile, ".")
        If Pos1 > 0 Then
        frmMain.lstDS.AddItem Left$(xFile, Pos1 - 1)
        Else
        frmMain.lstDS.AddItem xFile
        End If
          
        xFile = Dir
        DoEvents
    Loop

    'Load Script settings
    Call LoadDSSettings
    
    'Update frmmain.lblPlugsInstalled with Current number of Installed Scripts
    frmMain.lstDS.Visible = True 'Show scripts list
  
End Sub

'---------------------------------------------------------------------------------
'Purpose : Add backslash to script directory
'---------------------------------------------------------------------------------

Function FixPath(lPath As String) As String
    If Right(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

'---------------------------------------------------------------------------------
'Purpose : Main run point
'---------------------------------------------------------------------------------

Public Sub DebotnetRun()

On Error Resume Next

    Dim count As Integer
    Dim idx As Integer
    Dim PlugCnt As Integer

    'Cosmetics
    frmMain.lblBack.Enabled = False
    
    frmMain.lblRun.Enabled = False
    frmMain.lblRunSelected.Enabled = False
    frmMain.lblTestSelected.Enabled = False
    frmMain.lblUndoSelected.Enabled = False

    For count = 0 To frmMain.lstDS.ListCount - 1
        'Check if the Script is selected for processing.
        If (frmMain.lstDS.Selected(count)) Then
            'Script index.
            idx = (count + 1)
            With m_ScriptCls
                'Keeps a count of the number of Scripts selected.
                PlugCnt = (PlugCnt + 1)
                'Get Script filename.
                .FileName = m_ScriptCol(idx)
      
                'See if Script is loaded.
                If (.IsLoaded) Then
                    'Update status
                    frmMain.Caption = "Debotnet [" & .scID & " " & "Script " & idx & "/" & frmMain.lstDS.ListCount & "]"
                    'Select Active script
                    frmMain.lstDS.ListIndex = idx - 1
                    'Process the Script.
                    .PhasePlugRun
                     DoEvents
                    'Status
                    frmMain.Caption = "Debotnet"
                    
                  End If
            End With
        End If
   
    Next count
        
              
    'Some cosmetics again!
    
    PlugCnt = 0
    count = 0
    idx = 0

    frmMain.lblBack.Enabled = True
           
    frmMain.lblRun.Enabled = True
    frmMain.lblRunSelected.Enabled = True
    frmMain.lblTestSelected.Enabled = True
    frmMain.lblUndoSelected.Enabled = True

End Sub

'---------------------------------------------------------------------------------
'Purpose : Function to run SELECTED scripts in Debotnet ONLY
'---------------------------------------------------------------------------------

Public Sub DebotnetRunSelected()

On Error Resume Next

    Dim count As Integer
    Dim idx As Integer
    Dim PlugCnt As Integer

    'Cosmetics
    frmMain.lblBack.Enabled = False
        
    frmMain.lblRun.Enabled = False
    frmMain.lblRunSelected.Enabled = False
    frmMain.lblTestSelected.Enabled = False
    frmMain.lblUndoSelected.Enabled = False

            With m_ScriptCls
                'Get Script filename.
                .FileName = m_ScriptCol(idx)
      
                'See if Script is loaded.
                If (.IsLoaded) Then
                    'Process the Script.
                    .PhasePlugRun
                     DoEvents
                    'Status
                    'Call WriteActions("Script: " & .scID & " <completed>")
                    
                  End If
            End With

        
        PlugCnt = 0
        count = 0
        idx = 0
        
        'Cosmetics
        frmMain.lblBack.Enabled = True
                
        frmMain.lblRun.Enabled = True
        frmMain.lblRunSelected.Enabled = True
        frmMain.lblTestSelected.Enabled = True
        frmMain.lblUndoSelected.Enabled = True
     
End Sub

'---------------------------------------------------------------------------------
'Purpose : Function to TEST SELECTED scripts in Debotnet
'---------------------------------------------------------------------------------

Public Sub DebotnetTestSelected()

On Error Resume Next

    Dim count As Integer
    Dim idx As Integer
    Dim PlugCnt As Integer

    'Cosmetics
    frmMain.lblBack.Enabled = False
        
    frmMain.lblRun.Enabled = False
    frmMain.lblRunSelected.Enabled = False
    frmMain.lblTestSelected.Enabled = False
    frmMain.lblUndoSelected.Enabled = False

    Call WriteActions("SIMULATION ONLY >>>" & vbCrLf)
            With m_ScriptCls
                'Get Script filename.
                .FileName = m_ScriptCol(idx)
      
                'See if Script is loaded.
                If (.IsLoaded) Then
                    'Process the script.
                    .PhasePlugTest
                     DoEvents
                    'Status
                    'Call WriteActions("Script: " & .scID & " <completed>")
                  End If
            End With

        
        PlugCnt = 0
        count = 0
        idx = 0
        
        'Cosmetics
        frmMain.lblBack.Enabled = True
                
        frmMain.lblRun.Enabled = True
        frmMain.lblRunSelected.Enabled = True
        frmMain.lblTestSelected.Enabled = True
        frmMain.lblUndoSelected.Enabled = True
     
End Sub

'---------------------------------------------------------------------------------
'Purpose : Function to Undo SELECTED scripts in Debotnet ONLY
'---------------------------------------------------------------------------------

Public Sub DebotnetUndoSelected()

On Error Resume Next

    Dim count As Integer
    Dim idx As Integer
    Dim PlugCnt As Integer

    'Cosmetics
    frmMain.lblBack.Enabled = False
        
    frmMain.lblRun.Enabled = False
    frmMain.lblRunSelected.Enabled = False
    frmMain.lblTestSelected.Enabled = False
    frmMain.lblUndoSelected.Enabled = False

    
    For count = 0 To frmMain.lstDS.ListCount - 1
        'Check if the Script is selected for processing.
        If (frmMain.lstDS.Selected(count)) Then
            'Script index.
            idx = (count + 1)
            With m_ScriptCls
             'Keeps a count of the number of Scripts selected.
                PlugCnt = (PlugCnt + 1)
                'Get Script filename.
                .FileName = m_ScriptCol(idx)
      
                'See if Script is loaded.
                If (.IsLoaded) Then
                 'Update status
                    'frmMain.caption = "Script " & .scID & " " & "(Script " & idx & "/" & frmMain.lstDS.ListCount & ")"
                    'Select Active Script in Backup process
                    frmMain.lstDS.ListIndex = idx - 1
                    'Process the Script.
                    .PhasePlugUndo
                     DoEvents
                    'Status
                    'Call WriteActions("Script: " & .scID & " <completed>")
                    
                  End If
            End With
        End If
        
    Next count
    
        PlugCnt = 0
        count = 0
        idx = 0
        
        'Cosmetics
        frmMain.lblBack.Enabled = True
                
        frmMain.lblRun.Enabled = True
        frmMain.lblRunSelected.Enabled = True
        frmMain.lblTestSelected.Enabled = True
        frmMain.lblUndoSelected.Enabled = True
     
End Sub

'---------------------------------------------------------------------------------
'Purpose : Function to select scripts by Evaluation category
'---------------------------------------------------------------------------------

Public Function DebotnetSelect(Category As String) As String

On Error Resume Next

Dim count As Integer
Dim idx As Integer
Dim PlugCnt As Integer
    
Static X As Boolean
    
'Cosmetics
frmMain.lblBack.Enabled = False
        
frmMain.lblRun.Enabled = False
frmMain.lblRunSelected.Enabled = False
frmMain.lblTestSelected.Enabled = False
frmMain.lblUndoSelected.Enabled = False

    'Unselect all scripts
        X = (Not X)
        Call ItemsSelectAll(frmMain.lstDS, False)
    
         For count = 0 To frmMain.lstDS.ListCount - 1
    
            'Script index.
            idx = (count + 1)
            With m_ScriptCls
                'Keeps a count of the number of scripts selected.
                PlugCnt = (PlugCnt + 1)
                'Get Plug-in filename.
                .FileName = m_ScriptCol(idx)
                'See if Script is loaded.
                If (.IsLoaded) Then
                        
                    If .scEvaluation = Category Then
                    
                        frmMain.lstDS.ListIndex = idx - 1
                        
                        'Select script
                        frmMain.lstDS.Selected(frmMain.lstDS.ListIndex) = True
                        
                    End If
                        
                End If
                    
             End With
       
         Next count
        
        
        PlugCnt = 0
        count = 0
        idx = 0
        
        'Cosmetics
        frmMain.lblBack.Enabled = True
                
        frmMain.lblRun.Enabled = True
        frmMain.lblRunSelected.Enabled = True
        frmMain.lblTestSelected.Enabled = True
        frmMain.lblUndoSelected.Enabled = True
     
End Function

'---------------------------------------------------------------------------------
'Purpose : Extract filename from filepath // Used in script Import function \\
'---------------------------------------------------------------------------------

Public Function GetFilename(iPath As String) As String

   Dim II As Long
   
    If InStr(iPath, "\") = 0 Then
        GetFilename = iPath
    Else
        For II = Len(iPath) To 1 Step -1
            If Mid$(iPath, II, 1) = "\" Then
                GetFilename = Mid$(iPath, II + 1)
                Exit For
           End If
        Next II
    End If
    
End Function

' PIPELINE 0.8!
'---------------------------------------------------------------------------------
'Purpose  : Load command-line options
'           Supported parameter
'           /auto
'           /one
'           /custom
'---------------------------------------------------------------------------------

Public Sub LoadCLI()

On Error Resume Next 'Lazy and debug off
Dim count As Integer
Dim idx As Integer

Dim RunCLI As String

   If InStr(1, Command, "/auto") Then  'Debotnet runs silently and automatically ALL CATEGORIES AND SCRIPTS
   
        frmMain.Hide
        
            For count = 0 To frmMain.lstCS.ListCount - 1
                idx = (count + 1)
                frmMain.lstCS.ListIndex = idx - 1
    
                Call DebotnetRun     'Run Debotnet
                
            Next count
                
    
        MsgBox "Debotnet command-line mode has been successfully completed.", vbInformation, "CLI " & " /auto"
        End
            
    ElseIf InStr(1, Command, "/one") Then  'Debotnet runs silently and automatically ONLY SELECTED CATEGORY AND SCRIPTS
    
        frmMain.Hide
        
        Call DebotnetRun         'Run Debotnet
           
        MsgBox "Debotnet command-line mode has been successfully completed.", vbInformation, "CLI " & " /one"
        End
            
            
    ElseIf InStr(1, Command, "/custom") Then 'Debotnet loads a CUSTOM .SCRIPT PROFILE AND RUN ALL CATEGORIES AND SCRIPTS
       
        frmMain.Hide
                
        'Extract path and Show Output path after /custom
        RunCLI = Trim$(Mid$(Command, 4))
                
            'Remove doublequotes
            If Left$(RunCLI, 1) = """" Then
                RunCLI = Mid$(RunCLI, 2, Len(RunCLI) - 2)
            End If
                
                   
            'Load scripts
            For i = 0 To frmMain.lstDS.ListCount - 1
            
                frmMain.lstDS.Selected(i) = GetINIString(RunCLI, frmMain.lstCS, frmMain.lstDS.List(i))
            
            Next i
     
    
            'Run Debotnet over all categories
              For count = 0 To frmMain.lstCS.ListCount - 1
              
                idx = (count + 1)
                frmMain.lstCS.ListIndex = idx - 1
    
                Call DebotnetRun
                
            Next count
            
      
        MsgBox "Debotnet command-line mode has been successfully completed.", vbInformation, "CLI " & " /custom"
    
        End
            
    End If
    
End Sub

'---------------------------------------------------------------------------------
'Purpose  : Write log to frmmain.txtStatus
'---------------------------------------------------------------------------------

Public Sub WriteActions(ByVal sAction As String)

frmMain.txtStatus.Text = frmMain.txtStatus.Text & sAction & vbCrLf: frmMain.txtStatus.SelStart = Len(frmMain.txtStatus.Text)

End Sub

'---------------------------------------------------------------------------------
'Purpose : Select/unselect all spps in frmMain.lstDS
'---------------------------------------------------------------------------------

Public Sub ItemsSelectAll(ListB As ListBox, SellAll As Boolean)

On Error Resume Next

Dim cnt As Long
    For cnt = 0 To ListB.ListCount - 1
        ListB.Selected(cnt) = SellAll
    Next cnt
    
    ListB.Refresh
    ListB.ListIndex = 0
    
End Sub

'---------------------------------------------------------------------------------
'Purpose : Export scripts (Profile) to .SCRIPT file
'---------------------------------------------------------------------------------

Public Function ExportProfile()

On Error Resume Next

Dim count As Integer
Dim idx As Integer
Dim PlugCnt As Integer

Dim X As Integer
Dim i As Integer
Dim Export As String
    
Export = fncGetFileNametoSave("Profile (*.script)|*.script", "")
FileName$ = Export


    For count = 0 To frmMain.lstCS.ListCount - 1
        'Check if the Script is selected for processing.

            'Script index.
            idx = (count + 1)
            With m_ScriptCls
        
                'Keeps a count of the number of Scripts selected.
                PlugCnt = (PlugCnt + 1)
                'Get Script filename.
                .FileName = m_ScriptCol(idx)
        
                'See if Script is loaded.
                If (.IsLoaded) Then
                    'Select Active script
                    frmMain.lstCS.ListIndex = idx - 1
                    
                    For i = 0 To frmMain.lstDS.ListCount - 1
                      If frmMain.lstDS.Selected(i) Then
                        X = 1
                    Else
                        X = 0
                    End If
                        INI = WriteINI(FileName$, frmMain.lstCS, frmMain.lstDS.List(i), (X))
                    Next i
                    
                  End If

            End With

    Next count
        
        
        PlugCnt = 0
        count = 0
        idx = 0
        

End Function

'---------------------------------------------------------------------------------
'Purpose : Load scripts (Profile) from .SCRIPT File
'---------------------------------------------------------------------------------

Public Function ImportProfile()

On Error Resume Next
        

Dim count As Integer
Dim idx As Integer
Dim PlugCnt As Integer

Dim X As Integer
Dim i As Integer
Dim Import As String
    
Export = fncGetFileNametoSave("Profile (*.script)|*.script", "")
FileName$ = Export


    For count = 0 To frmMain.lstCS.ListCount - 1
        'Check if the Script is selected for processing.

            'Script index.
            idx = (count + 1)
            With m_ScriptCls
        
                'Keeps a count of the number of Scripts selected.
                PlugCnt = (PlugCnt + 1)
                'Get Script filename.
                .FileName = m_ScriptCol(idx)
        
                'See if Script is loaded.
                If (.IsLoaded) Then
                    'Select Active script
                    frmMain.lstCS.ListIndex = idx - 1
                    
                    
                        For i = 0 To frmMain.lstDS.ListCount - 1
                        
                            frmMain.lstDS.Selected(i) = GetINIString(FileName$, frmMain.lstCS, frmMain.lstDS.List(i))
                            frmMain.lstDS.ListIndex = 0
                        
                        Next i
                    
                                        
                        'Save Scripts Boolean Values to debotnet-settings.txt
                        Call SaveDSSettings
                        
                  End If
                  
            End With

   
    Next count
        
        
        PlugCnt = 0
        count = 0
        idx = 0

End Function

'---------------------------------------------------------------------------------
'Purpose : Export (selected) scripts to .TEXT file (E.g. for debugging purposes)
'---------------------------------------------------------------------------------

Public Function ExportScriptsToText()

On Error Resume Next

Dim count As Integer
Dim idx As Integer
Dim F As String
Dim Selected As Integer
Dim Header As String

Header = "Debotnet " & App.Major & "." & App.Minor & "." & App.Revision & " >>> [git:https://github.com/mirinsoft/debotnet]" & vbCrLf
                
  F = FreeFile
  
    Open fncGetFileNametoSave("Text Files (*.txt)|*.txt", "", "") For Output As #F
            
        'Print header
        Print #F, Header
        
        For count = 0 To frmMain.lstCS.ListCount - 1
        
           idx = (count + 1)
           frmMain.lstCS.ListIndex = idx - 1
            
            'Count total selected
            For t = 0 To frmMain.lstDS.ListCount - 1
                If frmMain.lstDS.Selected(t) = True Then Selected = Selected + 1
            Next t
            
           'Export selected
            For i = 0 To frmMain.lstDS.ListCount - 1
            
                If frmMain.lstDS.Selected(i) = True Then
                   Print #F, frmMain.lstCS.Text & " > " & frmMain.lstDS.List(i)
                 End If
                 
            Next i
        
        Next count
    
    'Print selected
    Print #F, "; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    Print #F, "In total " & Selected & " script(s) are selected."
    Print #F, Format(Now)
                   
    Close #F
    
End Function
