VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Debotnet"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15165
   BeginProperty Font 
      Name            =   "Segoe UI Semilight"
      Size            =   8.25
      Charset         =   0
      Weight          =   350
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   15165
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox PicTopLeft 
      BorderStyle     =   0  'Kein
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   3240
      TabIndex        =   15
      Top             =   0
      Width           =   3240
      Begin VB.Label lblAppName 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00000000&
         Caption         =   "d<botnet"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   15.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   360
         MouseIcon       =   "frmMain.frx":000C
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   16
         ToolTipText     =   "Click to open developer site"
         Top             =   200
         Width           =   1860
      End
   End
   Begin VB.PictureBox PicMiddle 
      BorderStyle     =   0  'Kein
      Height          =   8055
      Left            =   3240
      ScaleHeight     =   8055
      ScaleWidth      =   5745
      TabIndex        =   51
      Top             =   0
      Width           =   5750
      Begin VB.PictureBox PicFooterMiddle 
         BorderStyle     =   0  'Kein
         Height          =   360
         Left            =   -360
         ScaleHeight     =   360
         ScaleWidth      =   6120
         TabIndex        =   57
         Top             =   7680
         Width           =   6120
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   1080
            TabIndex        =   58
            Top             =   210
            Width           =   60
         End
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Text            =   "search"
         Top             =   235
         Visible         =   0   'False
         Width           =   5325
      End
      Begin VB.ListBox lstDS 
         Appearance      =   0  '2D
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   11.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   6960
         ItemData        =   "frmMain.frx":015E
         Left            =   240
         List            =   "frmMain.frx":0160
         Style           =   1  'Kontrollkästchen
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   960
         Width           =   5490
      End
      Begin VB.Label lblRightDivider 
         BackColor       =   &H00000000&
         ForeColor       =   &H00000000&
         Height          =   7680
         Left            =   5730
         TabIndex        =   56
         Top             =   0
         Width           =   15
      End
      Begin VB.Label lblLeftDivider 
         BackColor       =   &H00000000&
         ForeColor       =   &H00000000&
         Height          =   7800
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   15
      End
      Begin VB.Label lblScript 
         BackStyle       =   0  'Transparent
         Caption         =   "no script selected"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   54
         ToolTipText     =   "Version"
         Top             =   235
         Width           =   5325
      End
      Begin VB.Shape ShpSearch 
         BorderWidth     =   2
         Height          =   495
         Left            =   120
         Top             =   200
         Width           =   5500
      End
      Begin VB.Shape ShpSearchActive 
         BorderWidth     =   2
         Height          =   495
         Left            =   120
         Top             =   200
         Visible         =   0   'False
         Width           =   5500
      End
   End
   Begin VB.PictureBox PicSettings 
      BorderStyle     =   0  'Kein
      Height          =   8055
      Left            =   14880
      ScaleHeight     =   8055
      ScaleWidth      =   4455
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
      Begin VB.ListBox lstTheme 
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   12
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5070
         Left            =   360
         Sorted          =   -1  'True
         TabIndex        =   30
         ToolTipText     =   "Run always with high permissions, which allows to perform deeper system changes"
         Top             =   1680
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox txtWgetPath 
         Appearance      =   0  '2D
         Height          =   375
         Left            =   1680
         TabIndex        =   23
         Text            =   "WgetPath"
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtWgetParam 
         Appearance      =   0  '2D
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtRepository 
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   400
         Left            =   1320
         TabIndex        =   21
         Text            =   "https://github.com/mirinsoft/debotnet/blob/master/scripts/"
         Top             =   3720
         Width           =   2535
      End
      Begin VB.TextBox txtOutputDir 
         Appearance      =   0  '2D
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   400
         Left            =   1320
         TabIndex        =   18
         Text            =   "Select path"
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CheckBox chkUseDebotnetEditor 
         Appearance      =   0  '2D
         BackColor       =   &H80000004&
         Caption         =   "Internal Editor Integration"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   12
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   360
         TabIndex        =   12
         ToolTipText     =   "Switch between Debotnet's Script Editor and external (e.g. Notepad++)"
         Top             =   4440
         Value           =   1  'Aktiviert
         Width           =   3255
      End
      Begin VB.CheckBox chkRunAlwaysInElevatedMode 
         Appearance      =   0  '2D
         BackColor       =   &H80000004&
         Caption         =   "Run always as administrator"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   12
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   360
         TabIndex        =   50
         ToolTipText     =   "Run always with high permissions, which allows deeper system changes (e.g. in Registry > HKLM etc.)"
         Top             =   5040
         Width           =   3500
      End
      Begin VB.Label lblRepoURL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   9.75
            Charset         =   204
            Weight          =   350
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         MouseIcon       =   "frmMain.frx":0162
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   29
         Top             =   3765
         Width           =   465
      End
      Begin VB.Label lblRepoURLInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repository URL for updates"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   12
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   28
         ToolTipText     =   "Default repository used to update scripts from community"
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Label lblTheme 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change theme"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   12
            Charset         =   204
            Weight          =   350
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   360
         MouseIcon       =   "frmMain.frx":02B4
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   27
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label lblOutputDir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   9.75
            Charset         =   204
            Weight          =   350
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         MouseIcon       =   "frmMain.frx":0406
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   20
         Top             =   2400
         Width           =   585
      End
      Begin VB.Label lblOutputDirInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Output directory for downloads"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   12
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   19
         ToolTipText     =   "Some scripts can retrieve files from the web. Set here the output dir"
         Top             =   1800
         Width           =   3210
      End
      Begin VB.Label lblSettingsInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   360
         TabIndex        =   4
         Top             =   120
         Width           =   1320
      End
   End
   Begin VB.PictureBox PicFooterLeft 
      BorderStyle     =   0  'Kein
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   3240
      TabIndex        =   44
      Top             =   7680
      Width           =   3240
      Begin VB.Label lblAppVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "version"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   9
            Charset         =   204
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   1215
         MouseIcon       =   "frmMain.frx":0558
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   47
         ToolTipText     =   "Check for updates (Codename Pegasos)"
         Top             =   45
         Width           =   570
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00004000&
         Caption         =   "stable"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   360
         MouseIcon       =   "frmMain.frx":06AA
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   46
         ToolTipText     =   "Release channel"
         Top             =   45
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1080
         TabIndex        =   45
         Top             =   210
         Width           =   60
      End
   End
   Begin VB.PictureBox PicCode 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   9240
      ScaleHeight     =   4575
      ScaleWidth      =   5535
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox txtCode 
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4005
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   9
         Text            =   "frmMain.frx":07FC
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label lblCodeInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Press <ESC> to leave without saving"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.TextBox txtStatus 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   9480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   31
      Top             =   2400
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.ListBox lstCS 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   12.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4860
      ItemData        =   "frmMain.frx":0804
      Left            =   360
      List            =   "frmMain.frx":0806
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2880
   End
   Begin VB.PictureBox PicFooterRight 
      BorderStyle     =   0  'Kein
      Height          =   360
      Left            =   8950
      ScaleHeight     =   360
      ScaleWidth      =   5895
      TabIndex        =   5
      Top             =   7680
      Width           =   5895
      Begin VB.Label lblPatron 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "Patron "
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3480
         MouseIcon       =   "frmMain.frx":0808
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   59
         ToolTipText     =   "Patron"
         Top             =   45
         Width           =   735
      End
      Begin VB.Label lblImport 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Import"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   4320
         MouseIcon       =   "frmMain.frx":095A
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   43
         Top             =   45
         Width           =   555
      End
      Begin VB.Label lblCodeSymbol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   210
         Width           =   60
      End
      Begin VB.Label lblGitHub 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GitHub"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   204
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   5160
         MouseIcon       =   "frmMain.frx":0AAC
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   7
         ToolTipText     =   "Star Debotnet's repository on GitHub"
         Top             =   45
         Width           =   570
      End
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   9480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9240
      Top             =   720
   End
   Begin VB.PictureBox PicRight 
      BorderStyle     =   0  'Kein
      Height          =   735
      Left            =   9360
      ScaleHeight     =   735
      ScaleWidth      =   5655
      TabIndex        =   1
      Top             =   0
      Width           =   5655
      Begin VB.Label lblBack 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "< Back"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   11.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   49
         Top             =   270
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lblSaveScript 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   11.25
            Charset         =   204
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4035
         TabIndex        =   37
         ToolTipText     =   "Commit changes"
         Top             =   270
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblUpdateScript 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   11.25
            Charset         =   204
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1200
         TabIndex        =   36
         ToolTipText     =   "Update selected script(s) from GitHub"
         Top             =   270
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblReportScript 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   11.25
            Charset         =   204
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2235
         TabIndex        =   35
         ToolTipText     =   "Report issue with script on GitHub"
         Top             =   270
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblShareScript 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Share"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   11.25
            Charset         =   204
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3180
         TabIndex        =   34
         ToolTipText     =   "Tell your friend(s) about this script"
         Top             =   270
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblEditScript 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   11.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4035
         TabIndex        =   33
         ToolTipText     =   "View code of this script"
         Top             =   270
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lblDotMenu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ". . ."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4935
         TabIndex        =   32
         Top             =   255
         Width           =   300
      End
      Begin VB.Label lblHeaderSub 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   9.75
            Charset         =   204
            Weight          =   350
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         MouseIcon       =   "frmMain.frx":0BFE
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   2
         Top             =   525
         Width           =   60
      End
   End
   Begin VB.PictureBox PicDebug 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   615
      Left            =   9075
      ScaleHeight     =   615
      ScaleWidth      =   5775
      TabIndex        =   38
      Top             =   1650
      Width           =   5775
      Begin VB.Label lblRun 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Run"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   204
            Weight          =   600
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   400
         MouseIcon       =   "frmMain.frx":0D50
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   42
         Top             =   150
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label lblTestSelected 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test script"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   2280
         MouseIcon       =   "frmMain.frx":0EA2
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   41
         Top             =   240
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblRunSelected 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Run script"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   3250
         MouseIcon       =   "frmMain.frx":0FF4
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblUndoSelected 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Undo script"
         BeginProperty Font 
            Name            =   "Segoe UI Semilight"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   4200
         MouseIcon       =   "frmMain.frx":1146
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.PictureBox PicLeftNavMenu 
      BorderStyle     =   0  'Kein
      Height          =   6945
      Left            =   0
      ScaleHeight     =   6945
      ScaleWidth      =   3240
      TabIndex        =   0
      Top             =   720
      Width           =   3240
   End
   Begin VB.Label lblEditScriptRemote 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "View code on GitHub"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13680
      MouseIcon       =   "frmMain.frx":1298
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   48
      ToolTipText     =   "Edit code on GitHub"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label lblScriptDev 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FF00FF&
      Caption         =   "dev"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10335
      MouseIcon       =   "frmMain.frx":13EA
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblScriptVer 
      Alignment       =   2  'Zentriert
      BackColor       =   &H0000FFFF&
      Caption         =   "1.0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9480
      TabIndex        =   25
      ToolTipText     =   "Version"
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblScriptDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   24
      Top             =   840
      Width           =   60
   End
   Begin VB.Label lblEvaluation 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00008000&
      Caption         =   "Recommended"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11985
      TabIndex        =   14
      ToolTipText     =   "Evaluation"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblPackageDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   7800
      TabIndex        =   13
      Top             =   1485
      Width           =   45
   End
   Begin VB.Menu mnMain 
      Caption         =   "Main menu"
      Visible         =   0   'False
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuSettingsAdvanced 
         Caption         =   "Advanced settings"
      End
      Begin VB.Menu mnuRunAsAdmin 
         Caption         =   "Restart as Administrator"
      End
      Begin VB.Menu mnuDocs 
         Caption         =   "Documentation"
      End
      Begin VB.Menu mnuDevURL 
         Caption         =   "Visit the projects website"
      End
      Begin VB.Menu mnuUpdateCheck 
         Caption         =   "Check for updates"
      End
   End
   Begin VB.Menu mnContext 
      Caption         =   "Contex menu (txtStatus)"
      Visible         =   0   'False
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save to..."
      End
   End
   Begin VB.Menu mnTag 
      Caption         =   "Tag"
      Visible         =   0   'False
      Begin VB.Menu mnuAll 
         Caption         =   "tag all"
      End
      Begin VB.Menu mnuTagRecommended 
         Caption         =   "tag recommended"
      End
      Begin VB.Menu mnuTagLimited 
         Caption         =   "tag limited"
      End
      Begin VB.Menu mnuTagCustom 
         Caption         =   "tag custom"
      End
      Begin VB.Menu mnuTagBloatware 
         Caption         =   "tag bloatware"
      End
      Begin VB.Menu mnuUncheckAll 
         Caption         =   "tag none"
      End
   End
   Begin VB.Menu mnImport 
      Caption         =   "Import"
      Visible         =   0   'False
      Begin VB.Menu mnuImportScript 
         Caption         =   "Import script"
      End
      Begin VB.Menu mnuImportScriptRemote 
         Caption         =   "Import from Git"
      End
      Begin VB.Menu mnuImportProfile 
         Caption         =   "Import profile"
      End
      Begin VB.Menu mnuExportProfile 
         Caption         =   "Export profile"
      End
      Begin VB.Menu mnuExportTextFile 
         Caption         =   "Export to text file"
      End
      Begin VB.Menu mnuMarketplace 
         Caption         =   "Find more packs ..."
      End
   End
   Begin VB.Menu mnReleaseChannel 
      Caption         =   "Release channel"
      Visible         =   0   'False
      Begin VB.Menu mnuReleaseStable 
         Caption         =   "Stable"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuReleaseNightly 
         Caption         =   "Nightly"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------
'Description  :   Initial release 23-12-2019
'
'                 Debotnet is a tiny portable tool for controlling Windows 10's many privacy-related settings
'                 and keep your personal data private.
'
'                 Debotnet requires Windows 10 including both 32-bit and 64-bit versions.
'
'                 More infos can be found on
'                 https://github.com/mirinsoft/debotnet
'
'                 Copyright (c) 2019 Mirinsoft
'
'
'   WGET.EXE is not included in this sources. Download and put it to bin folder of Debotnet to update/download third-party script files!!!
'


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'Leave Code Editor with ESC
If PicCode.Visible = True Then

    If KeyCode = vbKeyEscape Then
        
        lblEditScript.Visible = True
                
        PicCode.Visible = False
        lblEditScriptRemote.Visible = False
        lblSaveScript.Visible = False
        
    End If
    
End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblEditScript.FontUnderline = False
lblSaveScript.FontUnderline = False
lblShareScript.FontUnderline = False
lblReportScript.FontUnderline = False
lblUpdateScript.FontUnderline = False

'Hide search
lblScript.Visible = True
ShpSearch.Visible = True

ShpSearchActive.Visible = False
txtSearch.Visible = False

End Sub

Private Sub lblImportRemote_Click()

    Msg = "Do you want to update the following script(s) based upon community optimizations from GitHub?"
    
        answer = MsgBox(Msg, vbInformation + vbYesNo, lblUpdateScript.Caption)
            
                  
        If answer = vbYes Then
            
      
            lblScriptDate.Caption = "Importing script [" & txtURLRemote.Text & "]"
            lblUpdateScript.Enabled = False
                    
            RetVal = ShellWait("cmd.exe /K" & """" & Chr(34) & txtWgetPath.Text & """" & " " & txtURLRemote.Text & "?raw=true" & "" & " " & "--show-progress --progress=bar:force --no-hsts --no-check-certificate --content-disposition -N -P" & " " & Chr(34) & App.Path & "\scripts\" & lstCS.Text & """", True)
        
       End If
    
        Call LoadDS   ' Refresh and load scripts
               
End Sub

Private Sub lblAppName_Click()

Call mnuDevURL_Click

End Sub

Private Sub lblDotMenu_Click()

 PopupMenu mnMain, 2
 
End Sub

Private Sub lblEditScript_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblEditScript.FontUnderline = True

End Sub

Private Sub lblBack_Click()

Call mnuClear_Click

'//Cosmetics

lblBack.Visible = False 'Hide Status

End Sub

Private Sub lblImport_Click()

PopupMenu mnImport, 2

End Sub

Private Sub lblImport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblImport.FontUnderline = True

End Sub

Private Sub lblImportRemoteCancel_Click()

PicImportRemote.Visible = False

End Sub

Private Sub lblPatron_Click()

Call lblScriptDev_Click
    
End Sub

Private Sub lblRepoURL_Click()

'Open repository

    Call ShellExecute(hwnd, "Open", _
        txtRepository.Text, "", "", 1)


End Sub

Private Sub lblTheme_Click()

    If lstTheme.Visible = True Then
    
        lstTheme.Visible = False
        
    Else
    
        lstTheme.Visible = True
    
    End If
    
End Sub

Private Sub lblUndoSelected_Click()


Dim UndoScripts As String

    For i = 0 To lstDS.ListCount - 1
    
        If lstDS.Selected(i) Then UndoScripts = UndoScripts & "- " & lstDS.List(i) & vbNewLine
        
    Next
    
 
If lstDS.SelCount >= 1 Then
    
        Msg = "Do you want to revert the changes in selected script(s)?" & vbNewLine & _
         "" & vbCrLf & _
        UndoScripts
            
        answer = MsgBox(Msg, vbExclamation + vbYesNo, lblUndoSelected.Caption)
                
                  
        If answer = vbYes Then
        
                frmMain.txtStatus.Text = "" 'Clear Status window
                txtStatus.Visible = True 'Show Status window
                lblBack.Visible = True 'Show Back button

                Call DebotnetUndoSelected
            
        End If
        
Else
      
    
    MsgBox "No scripts selected.", vbDefaultButton1
                   
    
End If
    

End Sub

Private Sub lstCS_GotFocus()

Call Settings_Close   'Close settings panel

End Sub

Private Sub lstTheme_Click()

Call LoadUI

End Sub

Private Sub lstDS_GotFocus()

Call Settings_Close   'Close settings panel

End Sub

Private Sub lstDS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
'Context-Menu mnTag
 
If Button = 2 Then

PopupMenu mnTag, 2

End If

End Sub

Private Sub mnuImportScriptRemote_Click()

Dim ans As String

    ans = InputBox("Paste here the URL of the script file from GitHub ", "Import script from GitHub", "https://www.github.com")
    
    If ans = "" Then
       'MsgBox "Import canceled"
    Else
   
              
        Me.Caption = "Importing script from Git [" & ans & "]"
                    
        RetVal = ShellWait("cmd.exe /c" & """" & Chr(34) & txtWgetPath.Text & """" & " " & ans & "?raw=true" & "" & " " & "--show-progress --progress=bar:force --no-hsts --no-check-certificate --content-disposition -N -P" & " " & Chr(34) & App.Path & "\scripts\" & lstCS.Text & """", False)
        
        Me.Caption = "Debotnet"
        
        Call LoadDS   ' Refresh and load scripts
        Call lstCS_Click  'Refresh category list
        
    End If
             

End Sub

Private Sub mnuMarketplace_Click()

'Open Marketplace site

   MsgBox "Not available anymore!"

End Sub

Public Sub mnuReleaseNightly_Click()

mnuReleaseNightly.Checked = Not mnuReleaseNightly.Checked

If mnuReleaseNightly.Checked = True Then
lblRelease.Caption = "nightly"

mnuReleaseStable.Checked = False

Else
lblRelease.Caption = "stable"
mnuReleaseStable.Checked = True
End If

'Refresh UI (color of release channel tag)
 Call LoadUI

End Sub

Private Sub mnuReleaseStable_Click()

mnuReleaseStable.Checked = Not mnuReleaseStable.Checked

If mnuReleaseStable.Checked = True Then
lblRelease.Caption = "stable"

mnuReleaseNightly.Checked = False

Else
lblRelease.Caption = "nightly"
mnuReleaseNightly.Checked = True
End If

'Refresh UI (Color of Release channel tag)
 Call LoadUI

End Sub

Private Sub mnuSave_Click()

'Export Status window log to text File

On Error Resume Next

Dim Header As String

Header = "Debotnet " & App.Major & "." & App.Minor & "." & App.Revision & " >>> [git:https://github.com/mirinsoft/debotnet]" & vbCrLf & vbCrLf
    
        FI = FreeFile
        Open fncGetFileNametoSave("Text Files (*.txt)|*.txt", "") For Output As #FI
        
            Print #FI, Header & txtStatus.Text
       
        Close FI
        
        
End Sub

Private Sub mnuExportTextFile_Click()

 Call ExportScriptsToText
        
End Sub

Private Sub mnuTagBloatware_Click()

'Tag bloatware

 Call DebotnetSelect("Bloatware")
 
End Sub

Private Sub mnuTagCustom_Click()

'Tag custom

 Call DebotnetSelect("Custom")
 
End Sub

Private Sub mnuTagLimited_Click()

'Tag limited

 Call DebotnetSelect("Limited")
 
End Sub

Private Sub PicFooterRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblImport.FontUnderline = False
lblGitHub.FontUnderline = False

End Sub

Private Sub PicLeftNavMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Save Scripts Boolean Values to debotnet-settings.txt
Call SaveDSSettings
End Sub

Private Sub PicLeftNavMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblAppVersion.FontUnderline = False

End Sub

Private Sub PicMiddle_Click()

Call Settings_Close   'Close settings panel

End Sub

Private Sub PicMiddle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Hide search
lblScript.Visible = True
ShpSearchActive.Visible = False

txtSearch.Visible = False
ShpSearch.Visible = True


End Sub

Private Sub txtStatus_GotFocus()

Call Settings_Close   'Close settings panel

End Sub

'---------------------------------------------------------------------------------
'Purpose : Select or Unselect all Items
'---------------------------------------------------------------------------------

Private Sub mnuAll_Click()

Static X As Boolean
    X = (Not X)
    
    Call ItemsSelectAll(lstDS, True)
    
End Sub

Private Sub mnuTagRecommended_Click()

'Tag recommended

 Call DebotnetSelect("Recommended")
 
End Sub

Private Sub mnuDocs_Click()

'Open documentation (GitHub)

    Call ShellExecute(hwnd, "Open", _
        "https://github.com/mirinsoft/debotnet/wiki", "", "", 1)
        
End Sub

Private Sub mnuExportProfile_Click()

Call ExportProfile

End Sub

Private Sub mnuImportProfile_Click()

 Call ImportProfile
 
End Sub

Private Sub mnuImportScript_Click()

On Error GoTo ErrHandler

Dim ImportScript As String

ImportScript = fncGetFileNametoOpen(, "Script (*.ds1)|*.ds1", "*.ds1") 'displays filename
    
    'Import Scripts first to default Scripts directory
    FileCopy ImportScript, App.Path & "\scripts\" & frmMain.lstCS.Text & "\" & GetFilename(ImportScript)
        
    'Refresh and Load apps/scripts list again
    lstDS.Clear
    Call LoadDS
 
      
    'Show Import Confirmation
    MsgBox GetINIString(ImportScript, "Info", "ID") & " by " & GetINIString(ImportScript, "Info", "Dev") & vbNewLine & _
    "Script has been successfully installed.", vbInformation, mnuImportScript.Caption
        
    'Delete Imported Script from Source then
    Kill ImportScript
    

ErrHandler:

    Exit Sub
    
End Sub

Private Sub mnuUncheckAll_Click()

'Unselect all in lstDS

    Call ItemsSelectAll(lstDS, False)
    
End Sub

Private Sub PicRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblEditScript.FontUnderline = False
lblSaveScript.FontUnderline = False
lblShareScript.FontUnderline = False
lblReportScript.FontUnderline = False
lblUpdateScript.FontUnderline = False

End Sub

Private Sub lblEditScriptRemote_Click()

   'Edit Script on GitHub
      Call ShellExecute(hwnd, "Open", _
        "https://github.com/mirinsoft/debotnet/edit/master/scripts/" & lstDS.Text & ".ds1", "", "", 1)
        

End Sub

Private Sub lblOutputDir_Click()

'Set Output directory for Wget
fNAME = BrowseForFolder(Me.hwnd, , "", True, True)
         
    If fNAME <> "" Then   'they did not hit cancel
         
        txtOutputDir.Text = fNAME
             
    End If
    
    
End Sub

Private Sub lblRelease_Click()
    
PopupMenu mnReleaseChannel, 2
        
End Sub

Private Sub lblRunSelected_Click()

'//Cosmetics
frmMain.txtStatus.Text = "" 'Clear Status window
txtStatus.Visible = True 'Show Status window
lblBack.Visible = True 'Show Back button
      
Call DebotnetRunSelected

End Sub

Private Sub lblScript_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Show search
ShpSearchActive.Visible = True
txtSearch.Visible = True

lblScript.Visible = False
ShpSearch.Visible = False

End Sub

Private Sub lblAppVersion_Click()

 
 If lblAppVersion.Caption = "Read release notes" Then
    
    'Show release notes on GitHub
      Call ShellExecute(hwnd, "Open", _
        "https://github.com/mirinsoft/debotnet/releases/tag/" & App.Major & "." & App.Minor & "." & App.Revision, "", "", 1)
    
  Else
        
     'Check for newer Version of Debotnet
        Call CheckAppUpdate
        
        
 End If
        
    
End Sub


Private Sub lblReportScript_Click()

    'Report Issue with script on GitHub
      Call ShellExecute(hwnd, "Open", _
        "https://github.com/mirinsoft/debotnet/issues/new", "", "", 1)

End Sub

Private Sub lblReportScript_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblReportScript.FontUnderline = True

End Sub

Private Sub lblSaveScript_Click()

On Error Resume Next

    With m_ScriptCls

        .FileName = m_ScriptCol(lstDS.ListIndex + 1)
        
        'Check if App is loaded.
        If (.IsLoaded) Then

        'Save with Debotnet Code Editor
        txt_WriteAll .FileName, txtCode.Text
        
        'Refresh
        Call lstDS_Click
        
        'Hide Code Editor
         PicCode.Visible = False
         
         lblEditScriptRemote.Visible = False
         lblSaveScript.Visible = False
         lblEditScript.Visible = True
            
      
        End If
    
    End With
    
    
End Sub

Private Sub lblSaveScript_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblSaveScript.FontUnderline = True

End Sub


Private Sub lblShareScript_Click()

On Error Resume Next

  Dim Result As Long
  Dim Buff As String
    
    
    Buff = "mailto:" & "@" & "?Subject=" & "Reclaim Windows 10 privacy with Debotnet [git:https://github.com/mirinsoft/debotnet]"
    Buff = Buff & "&Body=" & "I found this privacy script for Debotnet: " & lstDS.Text & ". You can get it here: https://github.com/mirinsoft/debotnet/tree/master/scripts/" & lstDS.Text & ".ds1"
    Result = ShellExecute(0&, "Open", Buff, "", "", 1)
    
    
End Sub

Private Sub lblShareScript_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblShareScript.FontUnderline = True

End Sub

Private Sub lblTestSelected_Click()

'//Cosmetics
frmMain.txtStatus.Text = "" 'Clear Status window
txtStatus.Visible = True 'Show Status window
lblBack.Visible = True 'Show Back button
            
Call DebotnetTestSelected

End Sub

Private Sub lblUpdateScript_Click()

Dim UpdateScripts As String

    For i = 0 To lstDS.ListCount - 1
    
        If lstDS.Selected(i) Then UpdateScripts = UpdateScripts & "- " & lstDS.List(i) & vbNewLine
        
    Next

        

If lstDS.SelCount >= 1 Then
    
    Msg = "Do you want to update the following script(s) based upon community optimizations from GitHub?" & vbNewLine & _
        "" & vbCrLf & _
         UpdateScripts
    
        answer = MsgBox(Msg, vbInformation + vbYesNo, lblUpdateScript.Caption)
            
                  
        If answer = vbYes Then
            
         
           For i = 0 To frmMain.lstDS.ListCount - 1
                    
                 If frmMain.lstDS.Selected(i) Then
            
                    lblScriptDate.Caption = "Updating script [" & lstDS.List(i) & "]"
                    lblUpdateScript.Enabled = False
                    
                
                    RetVal = ShellWait("cmd.exe /c" & """" & Chr(34) & txtWgetPath.Text & """" & " " & Chr(34) & txtRepository.Text & lstDS.List(i) & ".ds1?raw=true" & """" & " " & "--show-progress --progress=bar:force --no-hsts --no-check-certificate --content-disposition -N -P" & " " & Chr(34) & App.Path & "\scripts\" & lstCS.Text & """", False)
                    
                    'Wget Paramter -N enables time-stamping, which re-downloads the file only if its newer on the server than the downloaded version.
                    
                End If
                    
            Next i
            
                    lblUpdateScript.Enabled = True
                    Call lstDS_Click 'Refresh script list
                    
    
        End If
           
 Else
 
MsgBox "No scripts selected.", vbDefaultButton1, lblUpdateScript.Caption
        
End If

End Sub

Private Sub lblUpdateScript_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblUpdateScript.FontUnderline = True

End Sub

Private Sub lstCS_Click()

'Clear previous scripts in catetgory
lstDS.Clear

'Load scripts
Call LoadDS

'Refresh selected script(s)
Call lstDS_Click
lstDS.Refresh 'Refresh scripts list (white spots on dark themes)

End Sub

Private Sub lstDS_ItemCheck(Item As Integer)

On Error Resume Next
    
    'Show Warning (optional)
    With m_ScriptCls
        
        If lstDS.Visible = True Then
            
            .FileName = m_ScriptCol(lstDS.ListIndex + 1)
        
                If .scWarning <> "" Then
                    
                    If lstDS.Selected(lstDS.ListIndex) Then
                    MsgBox "WARNING about " & Chr(34) & .scID & Chr(34) & " script rule." & vbNewLine & _
                    "" & vbNewLine & _
                    .scWarning, vbExclamation
                
                End If
                    
            End If
                    
        End If
     
    End With

    Call SaveDSSettings  'Save Scripts Boolean Values to debotnet-settings.txt
    
End Sub

Private Sub txtStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

 'Context-Menu txtStatus
 
If Button = 2 Then

    txtStatus.Enabled = False
    PopupMenu mnContext, 2
    txtStatus.Enabled = True
    
End If

    
End Sub

Private Sub mnuClear_Click()

txtStatus.Text = ""
txtStatus.Visible = False
lblBack.Visible = False 'Show Back button

End Sub

Private Sub mnuDevURL_Click()

'Open website

    Call ShellExecute(hwnd, "Open", _
        "https://www.mirinsoft.com", "", "", 1)
            

End Sub

 Sub mnuRunAsAdmin_Click()

Dim Frm As Form
    
        'Close all forms and exit Debotnet
        For Each Frm In VB.Forms
              'Unload forms
              Unload Frm
              Set Frm = Nothing
        Next Frm
    
    'Run as Admin
    ShellExecute 0, "runas", App.Path & "\Debotnet.exe", Command, vbNullString, SW_SHOWNORMAL
  
  
End Sub

Private Sub mnuScriptImport_Click()

Call lblImport_Click

End Sub


Private Sub mnuSettings_Click()

PicSettings.Visible = True

End Sub

Private Sub mnuSettingsAdvanced_Click()

   DocumentOpen App.Path & "\bin\debotnet-settings.txt"
   
End Sub

Private Sub mnuUpdateCheck_Click()

 'Check for newer Version of Debotnet
    'Call CheckAppUpdate
    
    MsgBox "Add URL to version.txt in modUpdate > CheckAppUpdate() >  sURL ", vbCritical
End Sub

Private Sub Form_Activate()

    'Load Scripts and settings
    Call LoadDS
    
    'Load categories
    Call LoadCS
    
    'Preselect default category (from debotnet-settings.txt)
    frmMain.lstCS.ListIndex = GetINIString(App.Path & "\bin\debotnet-settings.txt", "Settings", "Category")
    
    'Load CLI/Command-line options
    Call LoadCLI
      
    'Show Release channel
    If frmMain.mnuReleaseStable.Checked = True Then
        frmMain.lblRelease.Caption = "stable"
    Else
        frmMain.lblRelease.Caption = "nightly"
    End If
        
End Sub

Private Sub Form_Click()

Call Settings_Close

End Sub

'---------------------------------------------------------------------------------
'Purpose : Saves Configuration files when exiting App
'---------------------------------------------------------------------------------

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save individual configuration settings
    Call SaveSettings
    
    'Save scripts to debotnet-settings.txt
    Call SaveDSSettings
        
    'Save controls states and positions
    Call SaveWindowState
    
    'Close all forms
    Dim Frm As Form
    'Loop through all forms
    For Each Frm In VB.Forms
          'Unload forms
          Unload Frm
          Set Frm = Nothing
    Next Frm
    
End Sub

Private Sub lblScriptDev_Click()

On Error Resume Next

    With m_ScriptCls
        
        .FileName = m_ScriptCol(lstDS.ListIndex + 1)
        'Check if Script is loaded
        If (.IsLoaded) Then
        

         Call ShellExecute(hwnd, "Open", _
            .scDevURL, "", "", 1)
    
        End If
        
    End With
    
End Sub

Private Sub lblScriptDev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    With m_ScriptCls
        
        'Check if App is loaded.
        If (.IsLoaded) Then
        
        'Open
        lblScriptDev.ToolTipText = .scDev & "," & .scDevURL
      
        End If
    
    End With
    
End Sub

Private Sub lblRun_Click()

Dim SelectedScripts As String

    For i = 0 To lstDS.ListCount - 1
    
        If lstDS.Selected(i) Then SelectedScripts = SelectedScripts & "- " & lstDS.List(i) & vbNewLine
        
    Next


If lstDS.SelCount >= 1 Then

 Msg = "Do you want to run selected scripts?" & vbNewLine & _
 "" & vbCrLf & _
 SelectedScripts
    
    answer = MsgBox(Msg, vbInformation + vbYesNo, lblRun.Caption)
        
              
        If answer = vbYes Then
        
            '//Cosmetics
            frmMain.txtStatus.Text = "" 'Clear Status window
            txtStatus.Visible = True 'Show Status window
            lblBack.Visible = True 'Show Back button
            
            Call DebotnetRun
        
        End If
        
    Else

        MsgBox "No scripts selected.", vbDefaultButton1
        
End If
        
End Sub

Private Sub lblEditScript_Click()

On Error Resume Next

    With m_ScriptCls

        .FileName = m_ScriptCol(lstDS.ListIndex + 1)
        
        'Check if App is loaded.
        If (.IsLoaded) Then
        
         If chkUseDebotnetEditor.Value = 1 Then        'Open with Debotnet Code editor
            
            PicCode.Visible = True
            txtCode.Text = txt_ReadAll(.FileName)
            
            'Show save button
            lblEditScriptRemote.Visible = True
            lblSaveScript.Visible = True
            lblEditScript.Visible = False
            
            
            Else      'Open with other Editor
          
            DocumentOpen .FileName
            
          End If
            
      
        End If
    
    End With

        
End Sub

Private Sub Form_Initialize()

'Remove all Listbox Borders
 Call RemoveListboxBorder
  
'Indicate the status of the current user and check whether it has Administrator privileges
 Call CheckAdministrator
 
'Add themes
 Call AddUI
 
'Load theme
 Call LoadUI

End Sub

Private Sub Form_Resize()

'  NO PRIOR YET! But stupidly solved!
On Error Resume Next

    'Controls
    PicMiddle.Height = Me.ScaleHeight
    PicRight.Left = Me.ScaleWidth - 5500
    PicLeftNavMenu.Height = Me.ScaleHeight

    'Divider
    lblLeftDivider.Height = Me.ScaleHeight 'Left
    lblRightDivider.Height = Me.ScaleHeight 'Right
    
    'Settings
    PicSettings.Left = Me.ScaleWidth - 4300
    PicSettings.Height = Me.ScaleHeight
    
    'Scripts
    lstDS.Height = PicMiddle.Height - 1100
    'Cats
    lstCS.Height = Me.ScaleHeight - 2000
     
    'Description
    txtDesc.Height = Me.ScaleHeight - 3000
    txtDesc.Width = Me.ScaleWidth - 9600

    'Status
    txtStatus.Width = Me.ScaleWidth - 9600
    txtStatus.Height = Me.ScaleHeight - 3000

    'Code Editor
    PicCode.Width = txtDesc.Width + 400
    PicCode.Height = Me.ScaleHeight - 2850
    txtCode.Height = PicCode.Height - 500
    txtCode.Width = PicCode.Width - 400
    lblEditScriptRemote.Left = Me.ScaleWidth - 2200
    
    'Footer
    PicFooterLeft.Top = Me.ScaleHeight + Me.ScaleTop - 350
    PicFooterMiddle.Top = Me.ScaleHeight + Me.ScaleTop - 350
    PicFooterRight.Top = Me.ScaleHeight + Me.ScaleTop - 350
    PicFooterRight.Width = Me.ScaleWidth
    lblPatron.Left = Me.ScaleWidth - 15500
    lblImport.Left = Me.ScaleWidth - 10700
    lblGitHub.Left = Me.ScaleWidth - 9800
    
End Sub

Private Sub lblGitHub_Click()

'GitHub

    Call ShellExecute(hwnd, "Open", _
        "https://github.com/mirinsoft/debotnet", "", "", 1)

End Sub

Private Sub lblGitHub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblGitHub.FontUnderline = True

End Sub


Private Sub lblAppVersion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblAppVersion.FontUnderline = True

End Sub

Private Sub lstDS_Click()

On Error Resume Next

Dim ScriptDate As String
    
    'Load Script info
    With m_ScriptCls
        
        .FileName = m_ScriptCol(lstDS.ListIndex + 1)
        'Check if Script is loaded
        If (.IsLoaded) Then
        
            '//Cosmetics
            lblScript.Enabled = True
            lblScriptVer.Visible = True
            lblScriptDev.Visible = True

            lblRun.Visible = True
            lblTestSelected.Visible = True
            lblRunSelected.Visible = True
            lblUndoSelected.Visible = True
            PicCode.Visible = False 'Close Code editor (if opened)
            lblEditScriptRemote.Visible = False 'Close Code on GitHub (if opened)
            
            lblEditScript.Visible = True
            lblSaveScript.Visible = False
            lblShareScript.Visible = True
            lblReportScript.Visible = True
            lblUpdateScript.Visible = True
            lblScriptDev.Enabled = True
            
            lblScript.Caption = .scID
            lblScriptVer.Caption = "v" & .scVer
            lblScriptDev.Caption = Left(.scDev, 17)
            txtDesc.Text = .scDesc
            
            'Show evaluation
            If .scEvaluation = "" Then
                lblEvaluation.Visible = False
            Else
                lblEvaluation.Visible = True
                lblEvaluation.Caption = .scEvaluation
                lblEvaluation.BackColor = HEXCOL2RGB(.scEvaluationColor)
            End If
            
            'Disable script updates
            If .scUpdate = "False" Then
                lblUpdateScript.Enabled = False
                lblEditScriptRemote.Enabled = False
            Else
                lblUpdateScript.Enabled = True
                lblEditScriptRemote.Enabled = True
            End If
            
            'Show Patron of the category
            If .scPatron = "" Then
                lblPatron.Visible = False
            Else
                lblPatron.Visible = True
                lblPatron.Caption = .scPatron
            End If
            
            'Show update/last file access
            ScriptDate = Format(FileDateTime(.FileName), "medium date")
            lblScriptDate.Caption = "last update " & ScriptDate
                     
        End If
    
    End With
    
    
Call Settings_Close   'Close settings panel
        
End Sub

Private Sub PicTopLeft_Click()

Call Settings_Close   'Close settings panel

End Sub

Private Sub PicLeftNavMenu_Click()

Call Settings_Close   'Close settings panel

End Sub

Private Sub PicRight_Click()

Call Settings_Close   'Close settings panel

End Sub

Private Sub PicSettings_Resize()

Call Settings_Close   'Close settings panel

End Sub

'---------------------------------------------------------------------------------
'Close Settings window
'---------------------------------------------------------------------------------

Private Sub Settings_Close()
    
PicSettings.Visible = False
lstTheme.Visible = False
    
End Sub

Private Sub Timer1_Timer()


    lblRun.Caption = "Run" & " (" & lstDS.SelCount & "/" & lstDS.ListCount & ")"
    
End Sub

Private Sub txtDesc_GotFocus()

Call Settings_Close   'Close settings panel

End Sub

'---------------------------------------------------------------------------------
'Purpose : Autocomplete in Search > frmmain.txtSearch
'---------------------------------------------------------------------------------

Private Sub txtSearch_Change()

Dim Pos As Long
    
lstDS.ListIndex = SendMessage(lstDS.hwnd, LB_FINDSTRING, -1, ByVal CStr(txtSearch.Text))

If lstDS.ListIndex = -1 Then
    Pos = txtSearch.SelStart
    
    Else
        
    Pos = txtSearch.SelStart
    txtSearch.Text = lstDS
    txtSearch.SelStart = Pos
    txtSearch.SelLength = Len(txtSearch.Text) - Pos
    
    End If
 
End Sub

Private Sub txtSearch_Click()

txtSearch.Text = ""

End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    
    If KeyCode = 8 Then 'Backspace
        If txtSearch.SelLength <> 0 Then
            txtSearch.Text = Mid$(txtSearch, 1, txtSearch.SelStart - 1)
            KeyCode = 0
        End If
    ElseIf KeyCode = 46 Then 'Del
        If txtSearch.SelLength <> 0 And _
            txtSearch.SelStart <> 0 Then
            txtSearch.Text = ""
            KeyCode = 0
        End If
    End If
    
End Sub

Private Sub txtSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

txtSearch.Text = lstDS.Text

End Sub
