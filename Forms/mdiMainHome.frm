VERSION 5.00
Begin VB.MDIForm mdiMainHome 
   BackColor       =   &H00000000&
   Caption         =   "Student  Attendance  Management  System"
   ClientHeight    =   9390
   ClientLeft      =   3240
   ClientTop       =   1815
   ClientWidth     =   15975
   Icon            =   "mdiMainHome.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox BackgroundCoverPicture 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   64800
      Left            =   0
      ScaleHeight     =   64800
      ScaleMode       =   0  'User
      ScaleWidth      =   22800
      TabIndex        =   0
      Top             =   0
      Width           =   22800
      Begin VB.CommandButton UserDashboardButton 
         BackColor       =   &H00C0C000&
         Caption         =   "User Dashboard"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   12360
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6240
         Width           =   3255
      End
      Begin VB.CommandButton AdminPanelButton 
         BackColor       =   &H00C0C000&
         Caption         =   "Admin  Panel"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6840
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6240
         Width           =   3135
      End
      Begin VB.Image UserDashboardImage 
         BorderStyle     =   1  'Fixed Single
         Height          =   1980
         Left            =   12960
         Picture         =   "mdiMainHome.frx":5A5A2
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   1980
      End
      Begin VB.Image AdminPanelImage 
         BorderStyle     =   1  'Fixed Single
         Height          =   1980
         Left            =   7440
         Picture         =   "mdiMainHome.frx":5B1CD
         Top             =   3840
         Width           =   1980
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Student  Attendance  Management  System"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   4800
         TabIndex        =   1
         Top             =   720
         Width           =   12975
      End
   End
   Begin VB.Menu HomeMenu 
      Caption         =   "  Home"
   End
   Begin VB.Menu AdminPanelMenu 
      Caption         =   "  Admin Panel"
   End
   Begin VB.Menu UserDashboardMenu 
      Caption         =   "  User Dashboard"
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "  Help"
      Begin VB.Menu DocumentationSubMenu 
         Caption         =   "Documentation"
      End
      Begin VB.Menu AboutSubMenu 
         Caption         =   "About"
      End
   End
   Begin VB.Menu ExitMenu 
      Caption         =   "  Exit"
   End
End
Attribute VB_Name = "mdiMainHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================== MAIN APPLICATION HUB =====================
' Central navigation interface for Student Attendance System
' Key Features:
'   - Role-based access (Admin/User)
'   - MDI child form management
'   - System documentation access
'   - Application exit control
' ===============================================================

Option Explicit

' --- ADMIN PANEL ACCESS (BUTTON) ---
' Initiates admin authentication sequence

Private Sub AdminPanelButton_Click()

    On Error Resume Next
        mdiMainHome.Enabled = False              ' Lock main MDI interface during login
        frmAdminPanel.isOpenAdminPanel = True    ' Set admin panel form open state flag
        frmAdminLogin.Show                       ' Display admin login form
        
End Sub

' --- USER DASHBOARD ACCESS (BUTTON) ---
' Opens user dashboard interface

Private Sub UserDashboardButton_Click()

    On Error Resume Next
        mdiMainHome.BackgroundCoverPicture.Visible = False  ' Hide default background of main MDI form
        frmUserDashboard.isOpenUserDashboard = True         ' Set user dashboard form open state flag
        frmUserDashboard.Show                               ' Launch user dashboard interface
        
End Sub

' --- HOME MENU SELECTION ---
' Returns application to default state

Private Sub HomeMenu_Click()

    closeMDIChildForms                          ' Close active child forms
    mdiMainHome.BackgroundCoverPicture.Visible = True  ' Show default background of main MDI form

End Sub

' --- ADMIN PANEL ACCESS (MENU) ---
' Alternate path to admin authentication

Private Sub AdminPanelMenu_Click()

    On Error Resume Next
        mdiMainHome.Enabled = False              ' Lock main MDI interface during login
        frmAdminPanel.isOpenAdminPanel = True    ' Set admin panel form open state flag
        frmAdminLogin.Show                       ' Display admin login form
        
End Sub

' --- USER DASHBOARD ACCESS (MENU) ---
' Alternate path to user dashboard

Private Sub UserDashboardMenu_Click()

    On Error Resume Next
        mdiMainHome.BackgroundCoverPicture.Visible = False  ' Hide default background of main MDI form
        frmUserDashboard.isOpenUserDashboard = True         ' Set user dashboard form open state flag
        frmUserDashboard.Show                               ' Launch user dashboard interface
        
End Sub

' --- DOCUMENTATION ACCESS ---
' Opens projects documentation in browser

Private Sub DocumentationSubMenu_Click()

    Dim pdfDocPath As String  ' Documentation storage path
    pdfDocPath = App.Path & "\Documents\Documentation (SAMS).pdf"  ' Build dynamic documentation path
    Shell "cmd /c start msedge """ & pdfDocPath & """", vbHide  ' Execute Edge silently with PDF
    
End Sub

' --- ABOUT SCREEN ACCESS ---
' Displays application information

Private Sub AboutSubMenu_Click()

    frmAbout.isOpenAbout = True  ' Set about form open state flag
    frmAbout.Show                ' Display about form
    
End Sub

' --- CHILD FORM CLEANUP ---
' Closes active admin/user modules

Private Sub closeMDIChildForms()

    ' Close all open child forms
    
    If frmAdminPanel.isOpenAdminPanel Then
        Unload frmAdminPanel     ' Terminate admin panel form
    End If
    
    If frmUserDashboard.isOpenUserDashboard Then
        Unload frmUserDashboard  ' Terminate user dashboard form
    End If
    
End Sub

' --- APPLICATION EXIT ---
' Safely terminates the application

Private Sub ExitMenu_Click()

    closeMDIChildForms          ' Cleanup active child forms
    
    ' Close about form if already open
    If frmAbout.isOpenAbout Then
        Unload frmAbout         ' Terminate about from
    End If

    Unload Me                   ' Terminate main MDI application
    
End Sub
