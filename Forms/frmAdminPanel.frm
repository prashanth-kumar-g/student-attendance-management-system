VERSION 5.00
Begin VB.Form frmAdminPanel 
   BackColor       =   &H00400040&
   Caption         =   "Admin Panel"
   ClientHeight    =   10080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   ControlBox      =   0   'False
   Icon            =   "frmAdminPanel.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   22800
   WindowState     =   2  'Maximized
   Begin VB.PictureBox MarkAttendancePicture 
      Height          =   1575
      Left            =   2520
      Picture         =   "frmAdminPanel.frx":5A5A2
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   15
      Top             =   3360
      Width           =   1575
   End
   Begin VB.PictureBox AddStudentPicture 
      Height          =   1635
      Left            =   8880
      Picture         =   "frmAdminPanel.frx":5CC8E
      ScaleHeight     =   1575
      ScaleWidth      =   1500
      TabIndex        =   14
      Top             =   3360
      Width           =   1560
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   6
      Left            =   15000
      TabIndex        =   11
      Top             =   3000
      Width           =   5295
      Begin VB.CommandButton EditStudentButton 
         BackColor       =   &H0000FF00&
         Caption         =   "EDIT STUDENT"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   2655
      End
      Begin VB.PictureBox EditStudentPicture 
         Height          =   1575
         Left            =   240
         Picture         =   "frmAdminPanel.frx":5FE37
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   4
      Left            =   12000
      TabIndex        =   8
      Top             =   6360
      Width           =   5295
      Begin VB.PictureBox EditAdminPicture 
         Height          =   1575
         Left            =   240
         Picture         =   "frmAdminPanel.frx":62EF6
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton EditAdminButton 
         BackColor       =   &H0000FF00&
         Caption         =   "EDIT ADMIN"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   2
      Left            =   5520
      TabIndex        =   5
      Top             =   6360
      Width           =   5295
      Begin VB.PictureBox AddAdminPicture 
         Height          =   1575
         Left            =   240
         Picture         =   "frmAdminPanel.frx":65A32
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton AddAdminButton 
         BackColor       =   &H0000FF00&
         Caption         =   "ADD ADMIN"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   1
      Left            =   8640
      TabIndex        =   3
      Top             =   3000
      Width           =   5295
      Begin VB.CommandButton AddStudentButton 
         BackColor       =   &H0000FF00&
         Caption         =   "ADD STUDENT"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   3000
      Width           =   5295
      Begin VB.CommandButton MarkAttendanceButton 
         BackColor       =   &H0000FF00&
         Caption         =   "MARK ATTENDANCE"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Image LogOutImage 
      Height          =   420
      Left            =   21120
      Picture         =   "frmAdminPanel.frx":6860D
      Stretch         =   -1  'True
      Top             =   840
      Width           =   420
   End
   Begin VB.Label LogOutLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   21720
      TabIndex        =   16
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Panel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9000
      TabIndex        =   0
      Top             =   720
      Width           =   4335
   End
End
Attribute VB_Name = "frmAdminPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================== ADMIN PANEL FORM =====================
' Main administrative interface for student attendance system
' Key Features:
'   - Central hub for administrative operations
'   - Sub form management (students/admins/attendance)
'   - Session management and logout handling
'   - State tracking for open forms
' ===========================================================

Option Explicit

' --- GLOBAL VARIABLES ---

Public isOpenAdminPanel As Boolean  ' Admin Panel form open state flag

' --- ATTENDANCE MANAGEMENT BUTTON CLICK ---
' Opens mark attendance form

Private Sub MarkAttendanceButton_Click()
    
    On Error Resume Next
        frmMarkAttendance.isOpenMarkAttendance = True  ' Set mark attendance form open state flag
        frmMarkAttendance.Show  ' Display mark attendance form

End Sub

' --- STUDENT MANAGEMENT: ADD BUTTON CLICK ---
' Opens add student form

Private Sub AddStudentButton_Click()
    
    On Error Resume Next
        frmAddStudent.isOpenAddStudent = True  ' Set add student form open state flag
        frmAddStudent.Show  ' Display add student form

End Sub

' --- STUDENT MANAGEMENT: EDIT BUTTON CLICK ---
' Opens edit student form

Private Sub EditStudentButton_Click()
    
    On Error Resume Next
        frmEditStudent.isOpenEditStudent = True  ' Set edit student form open state flag
        frmEditStudent.Show  ' Display edit student form

End Sub

' --- ADMIN MANAGEMENT: ADD BUTTON CLICK ---
' Opens add admin form

Private Sub AddAdminButton_Click()
    
    On Error Resume Next
        frmAddAdmin.isOpenAddAdmin = True  ' Set add admin form open state flag
        frmAddAdmin.Show  ' Display add admin form

End Sub

' --- ADMIN MANAGEMENT: EDIT BUTTON CLICK ---
' Opens edit admin form

Private Sub EditAdminButton_Click()
    
    On Error Resume Next
        frmEditAdmin.isOpenEditAdmin = True  ' Set edit admin form open state flag
        frmEditAdmin.Show  ' Display edit admin form

End Sub

' --- LOGOUT UI: IMAGE CLICK ---
' Triggers logout sequence when image clicked

Private Sub LogOutImage_Click()
    
    unloadAdminPanel  ' Initiate logout procedure

End Sub

' --- LOGOUT UI: LABEL CLICK ---
' Triggers logout sequence when label clicked

Private Sub LogOutLabel_Click()
    
    unloadAdminPanel  ' Initiate logout procedure

End Sub

' --- LOGOUT VISUAL SEQUENCE ---
' Provides visual feedback during logout

Private Sub unloadAdminPanel()
    
    LogOutLabel.Forecolor = &HC0&  ' Change label color to red
    DoEvents          ' Process pending events
    Sleep 1000        ' 1-second visual feedback delay
    Unload Me         ' Close admin panel interface

End Sub

' --- LOGOUT CLEANUP ROUTINE ---
' Cleans up open forms and resources

Private Sub logout()
    
    ' Close all open sub forms
    
    If frmMarkAttendance.isOpenMarkAttendance Then
        Unload frmMarkAttendance    ' Terminate mark attendance form
    End If

    If frmAddStudent.isOpenAddStudent Then
        Unload frmAddStudent    ' Terminate add student form
    End If
    
    If frmEditStudent.isOpenEditStudent Then
        Unload frmEditStudent   ' Terminate edit student form
    End If
    
    If frmAddAdmin.isOpenAddAdmin Then
        Unload frmAddAdmin      ' Terminate add admin form
    End If
    
    If frmEditAdmin.isOpenEditAdmin Then
        Unload frmEditAdmin     ' Terminate edit admin form
    End If
    
    ' Restore main MDI application interface
    mdiMainHome.BackgroundCoverPicture.Visible = True  ' Show default background of MDI form
    mdiMainHome.Show  ' Display main MDI interface
    
End Sub

' --- FORM UNLOAD EVENT ---
' Handles cleanup when form is closed

Private Sub Form_Unload(Cancel As Integer)
    
    isOpenAdminPanel = False  ' Reset Admin Details report open state flag
    
    logout  ' Execute cleanup routine

End Sub
