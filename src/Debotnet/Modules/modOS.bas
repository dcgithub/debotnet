Attribute VB_Name = "modOS"
'---------------------------------------------------------------------------------
'Public  Declaration  :  Indicate the status of the current user and check whether it has Administrator privileges
'---------------------------------------------------------------------------------

Public Declare Function IsUserAnAdmin Lib "shell32" Alias "#680" () As Integer

'---------------------------------------------------------------------------------
'Public  Declaration  :  Indicate the status of the current user and check whether it has Administrator privileges
'---------------------------------------------------------------------------------

Public Function CheckAdministrator()

      If IsUserAnAdmin() = 0 Then
                
                frmMain.Caption = "Debotnet"
                frmMain.mnuRunAsAdmin.Visible = True
                
            Else          ' Admin
            
                frmMain.Caption = "Debotnet (Administrator)"
                frmMain.mnuRunAsAdmin.Visible = False
            
                
            End If

End Function
