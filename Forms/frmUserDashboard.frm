VERSION 5.00
Begin VB.Form frmUserDashboard 
   BackColor       =   &H00404000&
   Caption         =   "User Dashboard"
   ClientHeight    =   9855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   ControlBox      =   0   'False
   Icon            =   "frmUserDashboard.frx":0000
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   9855
   ScaleWidth      =   22800
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   0
      Left            =   11880
      TabIndex        =   13
      Top             =   6360
      Width           =   5295
      Begin VB.PictureBox AdminDetailsPicture 
         Height          =   1575
         Left            =   240
         Picture         =   "frmUserDashboard.frx":5A5A2
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton AdminDetailsButton 
         BackColor       =   &H0000FF00&
         Caption         =   "ADMIN DETAILS"
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
         TabIndex        =   14
         Top             =   720
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   1
      Left            =   8640
      TabIndex        =   10
      Top             =   3000
      Width           =   5295
      Begin VB.CommandButton ViewAttendanceButton 
         BackColor       =   &H0000FF00&
         Caption         =   "VIEW ATTENDANCE"
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
         TabIndex        =   12
         Top             =   720
         Width           =   2655
      End
      Begin VB.PictureBox ViewAttendancePicture 
         Height          =   1575
         Left            =   240
         Picture         =   "frmUserDashboard.frx":5D0B1
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   0
      Left            =   2280
      TabIndex        =   7
      Top             =   3000
      Width           =   5295
      Begin VB.PictureBox IndividualAttendancePicture 
         Height          =   1575
         Left            =   240
         Picture         =   "frmUserDashboard.frx":5FA13
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton IndividualAttendanceButton 
         BackColor       =   &H0000FF00&
         Caption         =   "INDIVIDUAL ATTENDANCE"
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
         TabIndex        =   8
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   2
      Left            =   5520
      TabIndex        =   4
      Top             =   6360
      Width           =   5295
      Begin VB.CommandButton StudentDetailsButton 
         BackColor       =   &H0000FF00&
         Caption         =   "STUDENT DETAILS"
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
         Width           =   2775
      End
      Begin VB.PictureBox StudentDetailsPicture 
         Height          =   1575
         Left            =   240
         Picture         =   "frmUserDashboard.frx":620FF
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   4
      Left            =   15000
      TabIndex        =   1
      Top             =   3000
      Width           =   5295
      Begin VB.CommandButton OverallAttendanceButton 
         BackColor       =   &H0000FF00&
         Caption         =   "OVERALL ATTENDANCE"
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
         TabIndex        =   3
         Top             =   720
         Width           =   2655
      End
      Begin VB.PictureBox OverallAttendancePicture 
         Height          =   1575
         Left            =   240
         Picture         =   "frmUserDashboard.frx":65239
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label GoBackLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "GO BACK"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   21720
      TabIndex        =   16
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image GoBackImage 
      Height          =   465
      Left            =   21120
      Picture         =   "frmUserDashboard.frx":68447
      Stretch         =   -1  'True
      Top             =   840
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User Dashboard"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   9480
      TabIndex        =   0
      Top             =   720
      Width           =   4215
   End
End
Attribute VB_Name = "frmUserDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================== USER DASHBOARD FORM =====================
' Central interface for user-level operations in attendance system
' Key Features:
'   - Attendance tracking forms
'   - Database-driven report generation
'   - Visual navigation feedback
'   - Resource cleanup management
' ===============================================================

Option Explicit

' --- GLOBAL VARIABLES ---

Dim ConnDB As New ADODB.Connection          ' Database connection object
Public RsDataReport As New ADODB.Recordset  ' Database recordset object
Public isOpenUserDashboard As Boolean       ' User Dashboard form open state flag

' --- FORM LOAD EVENT ---
' Initializes database connection on form load

Private Sub Form_Load()
    
    On Error GoTo ConnectionError  ' Redirect runtime errors to ConnectionError handler
    
        ' Configure SQL Server database connection string
        ConnDB.ConnectionString = _
              "Provider = SQLOLEDB;" & _
              "Data Source = .\SQLEXPRESS02;" & _
              "Initial Catalog = SAMS;" & _
              "Integrated Security = SSPI;"
            
        ConnDB.Open   ' Establish database connection
        Exit Sub      ' Exit before error handler
    
' Database Connection Error Handler
ConnectionError:
        MsgBox "Database connection failed: " & Err.Description, vbCritical, "Error"    ' Show error details
        mdiMainHome.BackgroundCoverPicture.Visible = True  ' Show default background of MDI form
        mdiMainHome.Show  ' Display main MDI interface
        Unload Me         ' Terminate User Dashboard form
    
End Sub

' --- ATTENDANCE MODULE: INDIVIDUAL ATTENDANCE BUTTON CLICK ---
' Opens individual attendance form

Private Sub IndividualAttendanceButton_Click()
    
    On Error Resume Next
        frmIndividualAttendance.isOpenIndividualAttendance = True  ' Set individual attendance form open state flag
        frmIndividualAttendance.Show  ' Display individual attendance form

End Sub

' --- ATTENDANCE MODULE: VIEW ATTENDANCE BUTTON CLICK ---
' Opens view attendance form

Private Sub ViewAttendanceButton_Click()

    On Error Resume Next
        frmViewAttendance.isOpenViewAttendance = True  ' Set view attendance form open state flag
        frmViewAttendance.Show  ' Display view attendance form

End Sub

' (Note: To ensure data reports open correctly, set the printing preferences to Landscape and A4 size in the Control Panel.)

' --- ATTENDANCE MODULE: OVERALL ATTENDANCE BUTTON CLICK ---
' Generates overall attendance report

Private Sub OverallAttendanceButton_Click()

    ' Close overall attendance report if already open
    If isOpenOverallAttendanceDetails Then
        Unload dRepOverallAttendanceDetails  ' Terminate overall attendance report
    End If
        
    ' Generate overall attendance report
    RsDataReport.Open "SELECT * FROM overall_attendance_table", ConnDB, adOpenStatic, adLockReadOnly
        
    Set dRepOverallAttendanceDetails.DataSource = RsDataReport  ' Bind data to report
    dRepOverallAttendanceDetails.Show   ' Display overall attendance report
        
    isOpenOverallAttendanceDetails = True   ' Set overall attendance report open state flag

End Sub

' --- REPORTS MODULE: OVERALL ATTENDANCE LOST FOCUS ---
' Cleans up report resources

Private Sub OverallAttendanceButton_LostFocus()

    closeRsDataReport  ' Execute recordset cleanup

End Sub

' --- REPORTS MODULE: STUDENT DETAILS BUTTON CLICK ---
' Generates student details report

Private Sub StudentDetailsButton_Click()

    ' Close student details report if already open
    If isOpenStudentDetails Then
        Unload dRepStudentDetails   ' Terminate student details report
    End If
    
    ' Generate student details report
    RsDataReport.Open "SELECT * FROM student_table", ConnDB, adOpenStatic, adLockReadOnly
        
    Set dRepStudentDetails.DataSource = RsDataReport    ' Bind data to report
    dRepStudentDetails.Show     ' Display student details report
    
    isOpenStudentDetails = True     ' Set student details report open state flags

End Sub

' --- REPORTS MODULE: STUDENT DETAILS LOST FOCUS ---
' Cleans up report resources

Private Sub StudentDetailsButton_LostFocus()

    closeRsDataReport  ' Execute recordset cleanup

End Sub

' --- REPORTS MODULE: ADMIN DETAILS BUTTON CLICK ---
' Generates admin details report

Private Sub AdminDetailsButton_Click()

    ' Close admin details report if already open
    If isOpenAdminDetails Then
        Unload dRepAdminDetails     ' Terminate admin details report
    End If
    
    ' Generate admin details report
    RsDataReport.Open "SELECT * FROM admin_table", ConnDB, adOpenStatic, adLockReadOnly
        
    Set dRepAdminDetails.DataSource = RsDataReport    ' Bind data to report
    dRepAdminDetails.Show   ' Display admin details report
        
    isOpenAdminDetails = True   ' Set admin details report open state flag

End Sub

' --- REPORTS MODULE: ADMIN DETAILS LOST FOCUS ---
' Cleans up report resources

Private Sub AdminDetailsButton_LostFocus()

    closeRsDataReport  ' Execute recordset cleanup
    
End Sub

' --- RECORDSET CLEANUP ---
' Safely closes data report recordset

Private Sub closeRsDataReport()

    ' Close database recordset if open
    If RsDataReport.State = adStateOpen Then
        RsDataReport.Close     ' Close active recordset
        Set RsDataReport = Nothing  ' Release object memory
    End If

End Sub

' --- NAVIGATION UI: IMAGE CLICK ---
' Triggers logout sequence when image clicked

Private Sub GoBackImage_Click()

    unloadUserDashboard  ' Initiate navigation procedure

End Sub

' --- NAVIGATION UI: LABEL CLICK ---
' Triggers goback sequence when label clicked

Private Sub GoBackLabel_Click()

    unloadUserDashboard  ' Initiate navigation procedure

End Sub

' --- NAVIGATION VISUAL SEQUENCE ---
' Provides visual feedback during navigation

Private Sub unloadUserDashboard()

    GoBackLabel.Forecolor = &HC000&  ' Change label color to red
    DoEvents          ' Process pending events
    Sleep 1000        ' 1-second visual feedback delay
    Unload Me         ' Terminate user dashboard interface

End Sub

' --- NAVIGATION CLEANUP ROUTINE ---
' Cleans up open forms and resources

Private Sub goback()

    ' Close all open sub forms
    
    If frmIndividualAttendance.isOpenIndividualAttendance Then
        Unload frmIndividualAttendance  ' Terminate individual attendance form
    End If
    
    If frmViewAttendance.isOpenViewAttendance Then
        Unload frmViewAttendance  ' Terminate view attendance form
    End If
    
    If isOpenOverallAttendanceDetails Then
        Unload dRepOverallAttendanceDetails  ' Terminate overall attendance report
    End If
    
    If isOpenStudentDetails Then
        Unload dRepStudentDetails  ' Terminate student details report
    End If
    
    If isOpenAdminDetails Then
        Unload dRepAdminDetails  ' Terminate admin details report
    End If

    ' Restore main MDI application interface
    mdiMainHome.BackgroundCoverPicture.Visible = True  ' Show default background of MDI form
    mdiMainHome.Show  ' Display main MDI interface

End Sub

' --- FORM UNLOAD EVENT ---
' Handles cleanup when form is closed

Private Sub Form_Unload(Cancel As Integer)

    isOpenUserDashboard = False  ' Reset user dashboard open state flag

    goback  ' Execute cleanup routine

    ' Close database connection if open
    If ConnDB.State = adStateOpen Then
        ConnDB.Close    ' Close active connection
        Set ConnDB = Nothing    ' Release object memory
    End If

End Sub
