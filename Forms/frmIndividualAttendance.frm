VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIndividualAttendance 
   BackColor       =   &H00404000&
   Caption         =   "Individual Attendance"
   ClientHeight    =   10080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17010
   Icon            =   "frmIndividualAttendance.frx":0000
   LinkTopic       =   "Form11"
   ScaleHeight     =   10080
   ScaleWidth      =   17010
   WindowState     =   2  'Maximized
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00FF0000&
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox NameText 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   6240
      TabIndex        =   12
      Top             =   4800
      Width           =   3495
   End
   Begin VB.TextBox StudentIdText 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   6240
      TabIndex        =   11
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Frame AttendanceStatusFrame 
      BackColor       =   &H00808000&
      Caption         =   "Attendance Status"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   11760
      TabIndex        =   10
      Top             =   5760
      Width           =   5295
      Begin VB.Label AttendanceStatusLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   27.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   1440
         TabIndex        =   18
         Top             =   840
         Width           =   3975
      End
   End
   Begin VB.CommandButton CloseButton 
      BackColor       =   &H000000FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9720
      Width           =   2295
   End
   Begin VB.CommandButton NextButton 
      BackColor       =   &H0000FF00&
      Caption         =   "Next"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9720
      Width           =   2295
   End
   Begin VB.CommandButton FirstButton 
      BackColor       =   &H0000FF00&
      Caption         =   "First"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9720
      Width           =   2295
   End
   Begin VB.CommandButton PreviousButton 
      BackColor       =   &H0000FF00&
      Caption         =   "Previous"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9720
      Width           =   2295
   End
   Begin VB.CommandButton LastButton 
      BackColor       =   &H0000FF00&
      Caption         =   "Last"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9720
      Width           =   2295
   End
   Begin VB.TextBox CourseText 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   13440
      TabIndex        =   4
      Top             =   3840
      Width           =   3495
   End
   Begin VB.TextBox ClassText 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   13440
      TabIndex        =   3
      Top             =   4800
      Width           =   3495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      Format          =   134152193
      CurrentDate     =   44986
      MaxDate         =   73050
      MinDate         =   36526
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   9840
      TabIndex        =   17
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   9720
      TabIndex        =   16
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   2520
      TabIndex        =   15
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   2520
      TabIndex        =   14
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Image StudentImage 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   2295
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Students Photo"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   2520
      TabIndex        =   13
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Date"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   7
      Left            =   6120
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label0 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Individual Attendance"
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
      Index           =   1
      Left            =   7920
      TabIndex        =   0
      Top             =   360
      Width           =   7575
   End
End
Attribute VB_Name = "frmIndividualAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================== INDIVIDUAL ATTENDANCE FORM =====================
' Interface for viewing individual student attendance records
' Key Features:
'   - Date-based attendance navigation
'   - Student record browsing
'   - Visual status indicators
'   - Database resource management
' =====================================================================

Option Explicit

' --- GLOBAL VARIABLES ---

Dim ConnDB As New ADODB.Connection              ' Database connection object
Dim RsIndividualAttendance As New ADODB.Recordset  ' Attendance data recordset object
Public isOpenIndividualAttendance As Boolean   ' Individual Attendance form open state flag

' --- FORM LOAD EVENT ---
' Initializes date picker and database connection

Private Sub Form_Load()

    DTPicker1.Value = Date  ' Set default date to current date
    DTPicker1.Value = ""    ' Clear date selection
    
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
        MsgBox "Database openConn failed: " & Err.Description, vbCritical, "Error"    ' Show error details
        Unload Me   ' Terminate individual attendance form

End Sub

' --- DATE PICKER CHANGE EVENT ---
' Prepares interface when date changes

Private Sub DTPicker1_Change()
    
    disableControl    ' Disable navigation controls
    OKButton.Enabled = True  ' Enable attendance retrieval button
    clearControl      ' Clear student data display

End Sub

' --- OK BUTTON CLICK ---
' Retrieves attendance records for selected date

Private Sub OKButton_Click()
    
    ' Validate date selection
    If DTPicker1.Value = "" Then
        MsgBox "Select Date", vbExclamation, "Warning"  ' Show warning
        Exit Sub            ' Abort operation
    End If
        
    closeRsIndividualAttendance  ' Cleanup previous recordsets
        
    Dim sqlQuery As String  ' SQL query builder

    ' Build student attendance query
    sqlQuery = "SELECT student_table.Student_Id, student_table.Name, student_table.Photo, " & _
               "student_table.Course, student_table.Class, attendance_table.Attendance " & _
               "FROM student_table, attendance_table " & _
               "WHERE attendance_table.Date = '" & DTPicker1.Value & "' " & _
               "AND student_table.Student_Id = attendance_table.Student_Id " & _
               "ORDER BY student_table.Student_Id"
          
    ' Execute student details and attendance query
    RsIndividualAttendance.Open sqlQuery, ConnDB, adOpenKeyset, adLockReadOnly
    
    enableControl  ' Enable navigation controls after data retrieval
    
    OKButton.Enabled = False  ' Disable retrieval button after execution

End Sub

' --- FIRST RECORD BUTTON CLICK ---
' Navigates to first attendance record

Private Sub FirstButton_Click()

    ' Validate attendance data from helper function
    If Not hasAttendanceTaken() Then
        Exit Sub    ' if False then exit operation
    End If
    
    RsIndividualAttendance.MoveFirst  ' Move to first record
    loadValue  ' Display student data with their attendance status

End Sub

' --- NEXT RECORD BUTTON CLICK ---
' Navigates to next attendance record

Private Sub NextButton_Click()

   ' Validate attendance data from helper function
    If Not hasAttendanceTaken() Then
        Exit Sub    ' if False then exit operation
    End If
    
    RsIndividualAttendance.MoveNext  ' Move to next record
    
    ' Handle end of recordset
    If RsIndividualAttendance.EOF = True Then
        RsIndividualAttendance.MoveFirst  ' Wrap to first record
    End If
    
    loadValue  ' Display student data with their attendance status
    
End Sub

' --- PREVIOUS RECORD BUTTON CLICK ---
' Navigates to previous attendance record

Private Sub PreviousButton_Click()

    ' Validate attendance data from helper function
    If Not hasAttendanceTaken() Then
        Exit Sub    ' if False then exit operation
    End If
    
    RsIndividualAttendance.MovePrevious  ' Move to previous record
    
    ' Handle beginning of recordset
    If RsIndividualAttendance.BOF = True Then
       RsIndividualAttendance.MoveLast  ' Wrap to last record
    End If
    
    loadValue  ' Display student data with their attendance status

End Sub

' --- LAST RECORD BUTTON CLICK ---
' Navigates to last attendance record

Private Sub LastButton_Click()

    ' Validate attendance data from helper function
    If Not hasAttendanceTaken() Then
        Exit Sub    ' if False then exit operation
    End If
    
    RsIndividualAttendance.MoveLast  ' Move to last record
    loadValue  ' Display student data with their attendance status

End Sub

' --- CLOSE BUTTON CLICK ---
' Terminates Individual Attendance form

Private Sub CloseButton_Click()

    Unload Me  ' Terminate individual attendance interface

End Sub

' --- ATTENDANCE DATA VALIDATION ---
' Checks if attendance records exist for selected date
' Returns: True/False for record count

Private Function hasAttendanceTaken() As Boolean

    hasAttendanceTaken = (RsIndividualAttendance.RecordCount > 0)  ' Check record count and returns True/False

    ' Handle no attendance data
    If Not hasAttendanceTaken Then
        clearControl  ' Clear student data display
        MsgBox "No Attendance was taken on " & DTPicker1.Value, vbExclamation, "Warning"  ' Show warning
    End If

End Function

' --- NAVIGATION CONTROL ENABLE ---
' Activates record navigation buttons

Private Sub enableControl()

    FirstButton.Enabled = True     ' Enable first record button
    NextButton.Enabled = True      ' Enable next record button
    PreviousButton.Enabled = True  ' Enable previous record button
    LastButton.Enabled = True      ' Enable last record button

End Sub

' --- NAVIGATION CONTROL DISABLE ---
' Deactivates record navigation buttons

Private Sub disableControl()

    FirstButton.Enabled = False     ' Disable first record button
    NextButton.Enabled = False      ' Disable next record button
    PreviousButton.Enabled = False  ' Disable previous record button
    LastButton.Enabled = False      ' Disable last record button

End Sub

' --- RECORD DATA LOAD ---
' Populates UI with current record values

Private Sub loadValue()

    ' Display student information
    StudentIdText.Text = RsIndividualAttendance.Fields("Student_Id").Value  ' Set student ID
    NameText.Text = RsIndividualAttendance.Fields("Name").Value             ' Set student name
    StudentImage.Picture = LoadPicture(App.Path & "\" & RsIndividualAttendance.Fields("Photo").Value)  ' Set student photo from application path
    CourseText.Text = RsIndividualAttendance.Fields("Course").Value        ' Set course
    ClassText.Text = RsIndividualAttendance.Fields("Class").Value          ' Set class
    
    ' Display and color-code attendance status
    AttendanceStatusLabel.Caption = RsIndividualAttendance.Fields("Attendance").Value  ' Set attendance status
    If AttendanceStatusLabel.Caption = "Present" Then
        AttendanceStatusLabel.Forecolor = &HFF00&   ' Green color for present
    Else
        AttendanceStatusLabel.Forecolor = &HFF&     ' Red color for absent
    End If

End Sub

' --- UI CONTROL CLEAR ---
' Resets student data display

Private Sub clearControl()

    StudentIdText.Text = ""          ' Clear student ID
    NameText.Text = ""               ' Clear student name
    StudentImage.Picture = LoadPicture("")  ' Clear student photo
    CourseText.Text = ""             ' Clear course
    ClassText.Text = ""              ' Clear class
    AttendanceStatusLabel.Caption = ""  ' Clear attendance status

End Sub

' --- RECORDSET CLEANUP ---
' Safely closes attendance recordset

Private Sub closeRsIndividualAttendance()
    
    ' Close database recordset if open
    If RsIndividualAttendance.State = adStateOpen Then
        RsIndividualAttendance.Close     ' Close active recordset
        Set RsIndividualAttendance = Nothing  ' Release object memory
    End If
    
End Sub

' --- FORM UNLOAD EVENT ---
' Handles cleanup when form is closed

Private Sub Form_Unload(Cancel As Integer)

    isOpenIndividualAttendance = False  ' Reset individual attendance open state flag
    
    closeRsIndividualAttendance  ' Execute recordset cleanup
    
    ' Close database connection if open
    If ConnDB.State = adStateOpen Then
        ConnDB.Close    ' Close active connection
        Set ConnDB = Nothing    ' Release object memory
    End If

End Sub
