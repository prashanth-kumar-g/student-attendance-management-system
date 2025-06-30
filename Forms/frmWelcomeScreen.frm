VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWelcomeScreen 
   BackColor       =   &H00000000&
   Caption         =   "Welcome"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10815
   ControlBox      =   0   'False
   Icon            =   "frmWelcomeScreen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   10815
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   6000
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   105
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   10080
      Top             =   6120
   End
   Begin VB.Label ProgressBarLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8760
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Image WelcomeScreenImage 
      Height          =   3285
      Left            =   3840
      Picture         =   "frmWelcomeScreen.frx":5A5A2
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Student  Attendance Management System"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   -1080
      TabIndex        =   0
      Top             =   480
      Width           =   12975
   End
End
Attribute VB_Name = "frmWelcomeScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================== WELCOME SCREEN =====================
' Manages application startup with progress bar animation
' Key Features:
'   - Progress bar with custom coloring
'   - Auto-transition to main application window
'   - Percentage display during loading
' =========================================================

Option Explicit

' --- FORM LOAD EVENT ---
' Conditionally restores SAMS database if not present
' Note: Executes restoration script silently in background

Private Sub Form_Load()

    Dim BatchPath As String     ' Path to database restoration script
    Dim Command As String       ' Full command execution string
            
    ' Construct path to restoration batch file
    BatchPath = App.Path & "\Database\restore_sams.bat"
    
    ' Build execution command with proper shell formatting
    Command = "cmd.exe /C """ & BatchPath & """"
            
    ' Launch restoration process in hidden mode
    Shell Command, vbHide

End Sub

' --- TIMER TICK EVENT ---
' Handles progress bar increment and application initialization
' Note: Uses custom PBcolor module for progress bar styling

Private Sub Timer1_Timer()
    
    ' Suppress runtime errors during splash sequence
    On Error Resume Next
    
        ' Apply custom progress bar coloring (white background, green fill)
        Call PBcolor(ProgressBar1, vbWhite, vbGreen)
        
        ' Progress bar increment logic (5% per tick)
        If ProgressBar1.Value >= 0 Then
            ProgressBar1.Value = ProgressBar1.Value + 5     ' Increment progress bar by 5%
            ProgressBarLabel.Caption = ProgressBar1.Value & " %"    ' Update percentage display label
        End If
        
        ' When progress completes (exceeds 100%):
        If ProgressBar1.Value > 100 Then
            mdiMainHome.Show      ' Launch main MDI interface
            Unload Me             ' Terminate welcome screen form
        End If
        
End Sub
