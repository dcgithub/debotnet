Attribute VB_Name = "modProgram"

'---------------------------------------------------------------------------------
' Purpose   : Main entry point for the app
'---------------------------------------------------------------------------------

Public Sub Main()

On Error Resume Next 'Lazy and debug off (if theme file not available etc.)!!!
   
    Dim RunOnce As String, DateRun As String

        'Show Version
        frmMain.lblAppVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
         
        'Read Ini file
        FileName$ = App.Path
        
        If Right(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
        FileName$ = FileName$ & "\bin\debotnet-settings.txt"
        
        RunOnce = GetINIString(FileName$, "RunOnce", "RunOnce")
        DateRun = Date 'First run date

        'Load controls states and positions
        Call LoadWindowState
        
    Select Case RunOnce 'Select if Run or not
    
        Case 1 '// HAS BEEN RUN
 
        'Load Settings from .INI
        Call LoadSettings
                     
        'Load default theme
        frmMain.lstTheme.ListIndex = GetINIString(FileName$, "Settings", "Design")
                     
        'Run always as Admin?
            If frmMain.chkRunAlwaysInElevatedMode.Value = 1 And IsUserAnAdmin() = 0 Then
                
                'Load categories
                Call LoadCS
                
                'Preselect default category (from debotnet-settings.txt)
                frmMain.lstCS.ListIndex = GetINIString(App.Path & "\bin\debotnet-settings.txt", "Settings", "Category")
    
                'Close all forms and previous instances
                Dim Frm As Form
                For Each Frm In VB.Forms
                        Unload Frm
                        Set Frm = Nothing
                Next Frm

                'Run as admin
                ShellExecute 0, "runas", App.Path & "\Debotnet.exe", Command, vbNullString, SW_SHOWNORMAL
            
                Else   'Run normal
                    
                'Load frmMain
                frmMain.Show

            End If
    
            
        Case Else '// FIRST RUN
    
            '//Welcome screen
            MsgBox "With Debotnet you can easily configure Windows 10 in a privacy-friendly way by running simple command-line and PowerShell scripts." & vbNewLine & _
            vbNewLine & _
            "To perform deeper system changes (disable services, make changes in system and registry etc.) you should run Debotnet always as administrator.", vbInformation, "Welcome to Debotnet"
             
             'Check for Nightly release
             If FileExists(App.Path & "\nightly.dat") Then
                Call frmMain.mnuReleaseNightly_Click
             End If
             
            'Show release notes (ONLY in stable release)
             If frmMain.mnuReleaseStable.Checked = True Then
                frmMain.lblAppVersion.Caption = "Read release notes"
             End If
               
            '//Load defaults options in debotnet-settings.txt
             X = WriteINI(FileName$, "RunOnce", "RunOnce", "1")  'Changes to has been run
             X = WriteINI(FileName$, "RunOnce", "DateRun", DateRun) 'Set first run date
             X = WriteINI(FileName$, "Settings", "Design", "0")  'Set theme to default
             X = WriteINI(FileName$, "Settings", "Category", "0")   'Set start category to "start"

             X = WriteINI(FileName$, "Repository", "URL", frmMain.txtRepository.Text) 'Default repository
             X = WriteINI(FileName$, "Wget", "OutputDir", frmMain.txtOutputDir.Text) 'Download directory
             X = WriteINI(FileName$, "Wget", "WgetPath", frmMain.txtWgetPath.Text) 'Path to Wget.exe
             X = WriteINI(FileName$, "Wget", "WgetParam", frmMain.txtWgetParam.Text) 'Wget parameter
             
             '//Load default options in GUI
             frmMain.Width = "14835"
             frmMain.txtOutputDir.Text = App.Path   'Default output for Wget
             frmMain.txtWgetPath.Text = App.Path & "\bin\" & "wget.exe" 'Path to Wget.exe
             frmMain.txtWgetParam.Text = "-q --show-progress --progress=bar:force --no-hsts --no-check-certificate -N -P" 'Wget parameter
             
             'Load frmMain
             frmMain.Show
   
             'Load default theme
             frmMain.lstTheme.Text = "Debotnet"

    End Select
    
End Sub
