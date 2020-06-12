Attribute VB_Name = "modUI"
'---------------------------------------------------------------------------------
' Private Declaration :  Load Program Icon with 256 colors dynamically to form
'---------------------------------------------------------------------------------

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const LR_LOADFROMFILE = &H10
Private Const WM_SETICON = &H80
Private Const IMAGE_ICON = &H1
Private Const ICON_SMALL = &H0
Private Const ICON_BIG = &H1

'---------------------------------------------------------------------------------
'Private  Declaration  : Remove ListBox Border
'---------------------------------------------------------------------------------

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Const GWL_STYLE = (-16)
Const WS_BORDER = &H800000

'---------------------------------------------------------------------------------
'Purpose:  Load Program Icon with 256 colors dynamically to form
'---------------------------------------------------------------------------------

Public Sub SetIconFromFile(ByVal hwnd As Long, FullFileName As String)
    Dim hIcon As Long
    hIcon = LoadImage(App.hInstance, FullFileName, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
    If hIcon = 0 Then Exit Sub
    SendMessageLong hwnd, WM_SETICON, ICON_BIG, hIcon
    SendMessageLong hwnd, WM_SETICON, ICON_SMALL, hIcon
End Sub

'---------------------------------------------------------------------------------
'Purpose  : Remove Listbox Borders
'---------------------------------------------------------------------------------

Public Function RemoveListboxBorder()

'lstDS > frmMain
SetWindowLong frmMain.lstDS.hwnd, GWL_STYLE, GetWindowLong(frmMain.lstDS.hwnd, GWL_STYLE) And (Not WS_BORDER)

'lstCS > frmMain
SetWindowLong frmMain.lstCS.hwnd, GWL_STYLE, GetWindowLong(frmMain.lstCS.hwnd, GWL_STYLE) And (Not WS_BORDER)

'lstTheme > frmMain
SetWindowLong frmMain.lstTheme.hwnd, GWL_STYLE, GetWindowLong(frmMain.lstTheme.hwnd, GWL_STYLE) And (Not WS_BORDER)

End Function

'---------------------------------------------------------------------------------
'Purpose  :  Used to Convert a hex color string to an rgb color
'---------------------------------------------------------------------------------

Public Function HEXCOL2RGB(ByVal HexColor As String) As String

Dim Red As String
Dim Green As String
Dim Blue As String
Dim Color As String

Color = Replace(HexColor, "#", "")
    'Here HexColor = "00FF1F"

Red = Val("&H" & Mid(HexColor, 1, 2))
    'The red value is now the long version of "00"

Green = Val("&H" & Mid(HexColor, 3, 2))
    'The red value is now the long version of "FF"

Blue = Val("&H" & Mid(HexColor, 5, 2))
    'The red value is now the long version of "1F"

HEXCOL2RGB = RGB(Red, Green, Blue)
    'The output is an RGB value

End Function

Public Function AddUI()

On Error Resume Next 'Lazy and debug Off!

Dim xFile As String
Dim Pos1 As Long

 xFile = Dir(App.Path & "\bin\" & "*.design", vbDirectory)
    
    Do Until (xFile = "")
        
        Pos1 = InStrRev(xFile, ".")
        If Pos1 > 0 Then
        frmMain.lstTheme.AddItem Left$(xFile, Pos1 - 1)
        Else
        frmMain.lstTheme.AddItem xFile
        End If
          
        xFile = Dir
        DoEvents
    Loop


End Function

'---------------------------------------------------------------------------------
'Purpose  : Load UI appearance
'---------------------------------------------------------------------------------

Public Function LoadUI()

Dim AppBadge As String
Dim AppBadgeFontColor As String
Dim BackColor As String
Dim FontColor As String
Dim FontColorLight As String
Dim NavTop As String
Dim NavTopFontColor As String
Dim NavTopLeft As String
Dim Search As String
Dim SearchActive As String
Dim NavLeftMenu As String
Dim NavLeftMenuFontColor As String
Dim NavMiddleMenu As String
Dim SearchFontColor As String
Dim DebugFontColor As String
Dim Divider As String
Dim Footer As String
Dim FooterFontColor As String
Dim Settings As String
Dim SettingsFontColor As String

Dim FileName As String
    
FileName = App.Path & "\bin\" & frmMain.lstTheme.Text & ".design"

'//UI
'Load Program Icon
SetIconFromFile frmMain.hwnd, App.Path & "\bin\" & "Debotnet.ico"

 'Main colors
    AppBadge = GetINIString(FileName, "Theme", "AppBadge")
    AppBadgeFontColor = GetINIString(FileName, "Theme", "AppBadgeFontColor")
    BackColor = GetINIString(FileName, "Theme", "BackColor")
    FontColor = GetINIString(FileName, "Theme", "FontColor")
    FontColorLight = GetINIString(FileName, "Theme", "FontColorLight")
    NavTop = GetINIString(FileName, "Theme", "NavTop")
    NavTopFontColor = GetINIString(FileName, "Theme", "NavTopFontColor")
    NavLeftMenu = GetINIString(FileName, "Theme", "NavLeftMenu")
    NavLeftMenuFontColor = GetINIString(FileName, "Theme", "NavLeftMenuFontColor")
    NavMiddleMenu = GetINIString(FileName, "Theme", "NavMiddleMenu")
    Divider = GetINIString(FileName, "Theme", "Divider")
    Footer = GetINIString(FileName, "Theme", "Footer")
    FooterFontColor = GetINIString(FileName, "Theme", "FooterFontColor")
    DebugFontColor = GetINIString(FileName, "Theme", "DebugFontColor")
    Settings = GetINIString(FileName, "Theme", "Settings")
    SettingsFontColor = GetINIString(FileName, "Theme", "SettingsFontColor")
    Search = GetINIString(FileName, "Theme", "Search")
    SearchActive = GetINIString(FileName, "Theme", "SearchActive")
    SearchFontColor = GetINIString(FileName, "Theme", "SearchFontColor")
    
    'Set frmMain > Colors
    
    With frmMain
    
        'AppBadge
        .lblAppName.BackColor = HEXCOL2RGB(AppBadge)
        .lblAppName.ForeColor = HEXCOL2RGB(AppBadgeFontColor)
        
        '//BackColor
        .BackColor = HEXCOL2RGB(BackColor)
        
        '//Menu
        .lblDotMenu.ForeColor = HEXCOL2RGB(NavTopFontColor)
        
        '//Search
        .txtSearch.BackColor = HEXCOL2RGB(NavMiddleMenu)
        .txtSearch.ForeColor = HEXCOL2RGB(SearchFontColor)
        .ShpSearch.BorderColor = HEXCOL2RGB(Search)
        .ShpSearchActive.BorderColor = HEXCOL2RGB(SearchActive)
            
        '//Controls
        .lstDS.BackColor = HEXCOL2RGB(NavMiddleMenu)
        .lstCS.BackColor = HEXCOL2RGB(NavLeftMenu)
        .txtStatus.BackColor = HEXCOL2RGB(BackColor)
        .txtDesc.BackColor = HEXCOL2RGB(BackColor)
        .PicCode.BackColor = HEXCOL2RGB(BackColor)
        .txtCode.BackColor = HEXCOL2RGB(BackColor)
        .PicDebug.BackColor = HEXCOL2RGB(BackColor)
        
        '//UI Navigation controls
        .PicTopLeft.BackColor = HEXCOL2RGB(NavLeftMenu)
        .PicLeftNavMenu.BackColor = HEXCOL2RGB(NavLeftMenu)
        .PicMiddle.BackColor = HEXCOL2RGB(NavMiddleMenu)
        .PicRight.BackColor = HEXCOL2RGB(BackColor)
        .PicFooterLeft.BackColor = HEXCOL2RGB(Footer)
        .PicFooterMiddle.BackColor = HEXCOL2RGB(Footer)
        .PicFooterRight.BackColor = HEXCOL2RGB(Footer)

        '//Divider
        If GetINIString(FileName, "Theme", "Divider") = "" Then
            .lblLeftDivider.Visible = False
            .lblRightDivider.Visible = False
        Else
            .lblLeftDivider.Visible = True
            .lblRightDivider.Visible = True
            .lblLeftDivider.BackColor = HEXCOL2RGB(Divider)
            .lblRightDivider.BackColor = HEXCOL2RGB(Divider)
        End If
        
        '//Footer
        If GetINIString(FileName, "Theme", "Footer") = "" Then
            .PicFooterLeft.BackColor = HEXCOL2RGB(NavLeftMenu)
            .PicFooterMiddle.Visible = False
            .PicFooterRight.BackColor = HEXCOL2RGB(BackColor)
        Else
            .PicFooterLeft.Visible = True
            .PicFooterMiddle.Visible = True
            .PicFooterRight.Visible = True
        End If

        '//FontColor
        .lstDS.ForeColor = HEXCOL2RGB(FontColor)
        .lstCS.ForeColor = HEXCOL2RGB(NavLeftMenuFontColor)

        .lblScriptDate.ForeColor = HEXCOL2RGB(FontColorLight)
        .txtDesc.ForeColor = HEXCOL2RGB(FontColor)
        .PicCode.ForeColor = HEXCOL2RGB(FontColor)
        .txtCode.ForeColor = HEXCOL2RGB(FontColor)
        .lblCodeInfo.ForeColor = HEXCOL2RGB(FontColorLight)
        .lblRun.ForeColor = HEXCOL2RGB(FontColorLight)
        .lblRunSelected.ForeColor = HEXCOL2RGB(FontColorLight)
        .lblTestSelected.ForeColor = HEXCOL2RGB(FontColorLight)
        .lblUndoSelected.ForeColor = HEXCOL2RGB(FontColorLight)
        .lblImport.ForeColor = HEXCOL2RGB(FontColorLight)
        .lblBack.ForeColor = HEXCOL2RGB(NavTopFontColor)
        .lblEditScript.ForeColor = HEXCOL2RGB(NavTopFontColor)
        .lblSaveScript.ForeColor = HEXCOL2RGB(NavTopFontColor)
        .lblShareScript.ForeColor = HEXCOL2RGB(NavTopFontColor)
        .lblReportScript.ForeColor = HEXCOL2RGB(NavTopFontColor)
        .lblUpdateScript.ForeColor = HEXCOL2RGB(NavTopFontColor)
                
        '//NavTopFontColor
        .lblScript.ForeColor = HEXCOL2RGB(SearchFontColor)
        
        '//Status window
        .txtStatus.ForeColor = HEXCOL2RGB(DebugFontColor)
        
        '//Footer FontColor
        .lblAppVersion.ForeColor = HEXCOL2RGB(FooterFontColor)
        .lblGitHub.ForeColor = HEXCOL2RGB(FooterFontColor)
                
        'Settings
        .PicSettings.BackColor = HEXCOL2RGB(Settings)
        .lstTheme.BackColor = HEXCOL2RGB(Settings)
        .txtOutputDir.BackColor = RGB(222, 228, 255)
        .txtRepository.BackColor = RGB(222, 228, 255)
        .chkUseDebotnetEditor.BackColor = HEXCOL2RGB(Settings)
        .chkRunAlwaysInElevatedMode.BackColor = HEXCOL2RGB(Settings)
        
        '//Settings FontColor
        .lblSettingsInfo.ForeColor = HEXCOL2RGB(SettingsFontColor)
        .lblTheme.ForeColor = HEXCOL2RGB(SettingsFontColor)
        .lstTheme.ForeColor = HEXCOL2RGB(SettingsFontColor)
        .lblOutputDirInfo.ForeColor = HEXCOL2RGB(SettingsFontColor)
        .lblOutputDir.ForeColor = HEXCOL2RGB(SettingsFontColor)
        .lblRepoURLInfo.ForeColor = HEXCOL2RGB(SettingsFontColor)
        .lblRepoURL.ForeColor = HEXCOL2RGB(SettingsFontColor)
        .chkUseDebotnetEditor.ForeColor = HEXCOL2RGB(SettingsFontColor)
        .chkRunAlwaysInElevatedMode.ForeColor = HEXCOL2RGB(SettingsFontColor)
        
        '//Tags
        .lblScriptVer.BackColor = RGB(248, 226, 53) 'Ver
        .lblScriptDev.BackColor = RGB(0, 79, 134) 'Dev
        .lblPatron.BackColor = RGB(253, 250, 147) 'Patron
        
        'Release channel
        If .mnuReleaseStable.Checked = True Then
            .lblRelease.BackColor = RGB(0, 134, 114) 'Stable
        Else
            .lblRelease.BackColor = RGB(92, 37, 49) 'Nightly
        End If
    
    End With

End Function
