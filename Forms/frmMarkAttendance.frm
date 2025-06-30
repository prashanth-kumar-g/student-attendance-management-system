VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMarkAttendance 
   BackColor       =   &H00400040&
   Caption         =   "Mark Attendance"
   ClientHeight    =   9960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17130
   Icon            =   "frmMarkAttendance.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   9960
   ScaleWidth      =   17130
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00FF0000&
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2280
      Width           =   735
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
      Left            =   12960
      TabIndex        =   17
      Top             =   4800
      Width           =   3495
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
      Left            =   12960
      TabIndex        =   16
      Top             =   3840
      Width           =   3495
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   15
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   14
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   13
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9720
      Width           =   2295
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
      Left            =   15600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9720
      Width           =   2295
   End
   Begin VB.Frame MarkAttendanceFrame 
      BackColor       =   &H00800080&
      Caption         =   "Mark Attendance"
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
      Left            =   11280
      TabIndex        =   8
      Top             =   5760
      Width           =   5295
      Begin VB.OptionButton AbsentOption 
         BackColor       =   &H00000000&
         Caption         =   "Absent"
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
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   2880
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton PresentOption 
         BackColor       =   &H00000000&
         Caption         =   "Present"
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
         ForeColor       =   &H0000FF00&
         Height          =   735
         Left            =   600
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label MarkedAttendanceLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Attendance Marked"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1095
         Left            =   1320
         TabIndex        =   19
         Top             =   1800
         Visible         =   0   'False
         Width           =   2775
      End
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
      Left            =   5760
      TabIndex        =   2
      Top             =   3840
      Width           =   3495
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
      Left            =   5760
      TabIndex        =   1
      Top             =   4800
      Width           =   3495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   5760
      TabIndex        =   20
      Top             =   2280
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
      Left            =   5640
      TabIndex        =   21
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label UpdateAttendanceLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Update Attendance"
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
      Height          =   615
      Left            =   16680
      TabIndex        =   18
      Top             =   6960
      Visible         =   0   'False
      Width           =   2295
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
      Left            =   2040
      TabIndex        =   7
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Image StudentImage 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   2295
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2175
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
      Left            =   2040
      TabIndex        =   6
      Top             =   3960
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
      Left            =   2040
      TabIndex        =   5
      Top             =   4920
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
      Left            =   9240
      TabIndex        =   4
      Top             =   4920
      Width           =   2895
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
      Left            =   9360
      TabIndex        =   3
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mark Attendance"
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
      Height          =   975
      Index           =   0
      Left            =   9000
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmMarkAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================== MARK ATTENDANCE FORM =====================
' Interface for recording and updating student attendance records
' Key Features:
'   - Date-based attendance tracking
'   - Student record navigation
'   - Present/Absent status marking
'   - Real-time attendance statistics calculation
'   - Database synchronization with error handling
' ===============================================================

Option Explicit

' --- GLOBAL VARIABLES ---

Dim ConnDB As New ADODB.Connection             ' Database connection object
Dim RsStudentDetails As New ADODB.Recordset    ' Student data recordset object
Dim RsAttendanceStatus As New ADODB.Recordset  ' Current attendance status recordset
Dim sqlQuery As String                         ' SQL command builder
Dim signalUpdateAttendance As Integer          ' Attendance update flag (0=New, 1=Update)
Public isOpenMarkAttendance As Boolean         ' Mark Attendance form open state flag

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
        MsgBox "Database connection failed: " & Err.Description, vbCritical, "Error"    ' Show error details
        Unload Me   ' Terminate mark attendance form

End Sub

' --- DATE PICKER CHANGE EVENT ---
' Prepares interface when date changesl

Private Sub DTPicker1_Change()
    
    disableControl     ' Deactivate navigation controls
    disableOption      ' Deactivate attendance options
    clearControl       ' Clear student data display
    OKButton.Enabled = True  ' Enable attendance retrieval button

End Sub

' --- OK BUTTON CLICK ---
' Retrieves student records for attendance marking

Private Sub OKButton_Click()
    
    ' Validate date selection
    If DTPicker1.Value = "" Then
        MsgBox "Select Date", vbExclamation, "Warning"  ' Show warning
        Exit Sub            ' Abort operation
    End If
        
    closeRsAttendanceDetails   ' Cleanup previous student recordsets
    closeRsAttendanceStatus    ' Cleanup previous attendance recordsets
        
    ' Open student details recordset
    RsStudentDetails.Open "SELECT Student_Id,Name,Photo,Course,Class FROM student_table", ConnDB, adOpenKeyset, adLockReadOnly

    ' Check if any student records exist
    If RsStudentDetails.EOF And RsStudentDetails.BOF Then
        MsgBox "No student records found in the database." & vbCrLf & _
               "Please add students before marking attendance.", _
               vbExclamation, "Warning"     ' Show warning
        Exit Sub    ' Exit if no student records
    End If

    ' Build current attendance status query
    sqlQuery = "SELECT Student_Id, Attendance FROM attendance_table WHERE Date='" & DTPicker1.Value & "' ORDER BY Student_Id"
    ' Execute attendance status query
    RsAttendanceStatus.Open sqlQuery, ConnDB, adOpenKeyset, adLockReadOnly
    
    enableControl  ' Activate navigation controls after data retrieval
    
    OKButton.Enabled = False  ' Disable retrieval button after execution

End Sub

' --- FIRST RECORD BUTTON CLICK ---
' Navigates to first student record

Private Sub FirstButton_Click()

    enableOption               ' Activate attendance Present/Absent options
    MarkAttendanceFrame.Enabled = True  ' Enable attendance marking frame
    UpdateAttendanceLabel.Forecolor = &HFFFFFF  ' Reset update label color to white
    hideLabel                  ' Hide attendance status labels
    
    RsStudentDetails.MoveFirst ' Move to first student record
    
    loadValue          ' Display student data
    isAttendanceMarked ' Check attendance status

End Sub

' --- NEXT RECORD BUTTON CLICK ---
' Navigates to next student record

Private Sub NextButton_Click()

    enableOption               ' Activate attendance Present/Absent options
    MarkAttendanceFrame.Enabled = True  ' Enable attendance marking frame
    UpdateAttendanceLabel.Forecolor = &HFFFFFF  ' Reset update label color to white
    hideLabel                  ' Hide attendance status labels
    
    RsStudentDetails.MoveNext  ' Move to next student record
    
    ' Handle end of recordset
    If RsStudentDetails.EOF = True Then
        RsStudentDetails.MoveFirst  ' Wrap to first student record
    End If
    
    loadValue          ' Display student data
    isAttendanceMarked ' Check if attendance already marked

End Sub

' --- PREVIOUS RECORD BUTTON CLICK ---
' Navigates to previous student record

Private Sub PreviousButton_Click()

    enableOption               ' Activate attendance Present/Absent options
    MarkAttendanceFrame.Enabled = True  ' Enable attendance marking frame
    UpdateAttendanceLabel.Forecolor = &HFFFFFF  ' Reset update label color to white
    hideLabel                  ' Hide attendance status labels
    
    RsStudentDetails.MovePrevious  ' Move to previous student record
    
    ' Handle beginning of recordset
    If RsStudentDetails.BOF = True Then
        RsStudentDetails.MoveLast  ' Wrap to last student record
    End If
    
    loadValue          ' Display student data
    isAttendanceMarked ' Check if attendance already marked

End Sub

' --- LAST RECORD BUTTON CLICK ---
' Navigates to last student record

Private Sub LastButton_Click()

    enableOption               ' Activate attendance Present/Absent options
    MarkAttendanceFrame.Enabled = True  ' Enable attendance marking frame
    UpdateAttendanceLabel.Forecolor = &HFFFFFF  ' Reset update label color to white
    hideLabel                  ' Hide attendance status labels
    
    RsStudentDetails.MoveLast  ' Move to last student record
    
    loadValue          ' Display student data
    isAttendanceMarked ' Check if attendance already marked
    
End Sub

' --- PRESENT OPTION CLICK ---
' Marks student as present for selected date

Private Sub PresentOption_Click()

    Dim AttendanceStatus, AttendanceUpdated, AttendanceMarked As String     ' Attendance status and confirmation messages
    
    AttendanceStatus = "Present"  ' Set attendance status to Present
    AttendanceUpdated = "Attendance Updated, Marked Present."  ' Update confirmation message
    AttendanceMarked = "Attendance Marked Present"  ' New record confirmation message
    
    ' Execute attendance marking procedure
    Call markAttendance(AttendanceStatus, AttendanceUpdated, AttendanceMarked)

End Sub

' --- ABSENT OPTION CLICK ---
' Marks student as absent for selected date

Private Sub AbsentOption_Click()

    Dim AttendanceStatus, AttendanceUpdated, AttendanceMarked As String     ' Attendance status and confirmation messages
    
    AttendanceStatus = "Absent"  ' Set attendance status to Absent
    AttendanceUpdated = "Attendance Updated, Marked Absent."  ' Update confirmation message
    AttendanceMarked = "Attendance Marked Absent"  ' New record confirmation message
    
    ' Execute attendance marking procedure
    Call markAttendance(AttendanceStatus, AttendanceUpdated, AttendanceMarked)

End Sub

' --- UPDATE ATTENDANCE LABEL CLICK ---
' Enables attendance modification for current student

Private Sub UpdateAttendanceLabel_Click()

    UpdateAttendanceLabel.Forecolor = &HFF&  ' Change label color to red
    MarkAttendanceFrame.Enabled = True       ' Enable attendance frame
    MarkedAttendanceLabel.Visible = False    ' Hide marked status label
    enableOption               ' Activate attendance options
    signalUpdateAttendance = 1 ' Set update flag

End Sub

' --- CLOSE BUTTON CLICK ---
' Terminates Mark Attendance form

Private Sub CloseButton_Click()

    Unload Me  ' Terminate mark attendance interface

End Sub

' --- RECORD DATA LOAD ---
' Populates UI with current student record values

Private Sub loadValue()

    StudentIdText.Text = RsStudentDetails.Fields("Student_Id").Value  ' Set student id field
    NameText.Text = RsStudentDetails.Fields("Name").Value             ' Set student name field
    StudentImage.Picture = LoadPicture(App.Path & "\" & RsStudentDetails.Fields("Photo").Value) ' Set photo field
    CourseText.Text = RsStudentDetails.Fields("Course").Value         ' Set course field
    ClassText.Text = RsStudentDetails.Fields("Class").Value           ' Set class field

End Sub

' --- ATTENDANCE STATUS CHECK ---
' Determines if attendance is already recorded for current student

Private Sub isAttendanceMarked()
    
    RsAttendanceStatus.Requery  ' Refresh attendance status data from database

    ' Handle empty recordset
    If RsAttendanceStatus.RecordCount = 0 Then
        Exit Sub    ' Exit if no attendance records
    End If

    RsAttendanceStatus.MoveFirst    ' Start from first attendance record
        
    ' Search for current student's attendance status
    Do Until RsAttendanceStatus.EOF
        ' Check if student is matches and attendance is marked
        If RsAttendanceStatus.Fields("Student_Id").Value = StudentIdText.Text And _
           (RsAttendanceStatus.Fields("Attendance").Value = "Present" Or RsAttendanceStatus.Fields("Attendance").Value = "Absent") Then
                disableOption       ' Deactivate attendance options
                visibleLabel        ' Show status labels
                Exit Sub            ' Exit after finding record
        End If
        RsAttendanceStatus.MoveNext    ' Move to next attendance record
    Loop

End Sub


' --- ATTENDANCE RECORDING PROCEDURE ---
' Saves or updates attendance record in database
' Calculates and updates overall attendance statistics
' Parameters:
'   AttendanceStatus: "Present" or "Absent"
'   AttendanceUpdated: Update confirmation message
'   AttendanceMarked: New record confirmation message

Private Sub markAttendance(ByVal AttendanceStatus As String, ByVal AttendanceUpdated As String, ByVal AttendanceMarked As String)

    On Error GoTo MarkAttendanceError   ' Setup error handling

        If signalUpdateAttendance = 1 Then    ' Update existing attendance record
            ' Build attendance update query
            sqlQuery = "UPDATE attendance_table SET Attendance = '" & AttendanceStatus & "' " & _
                       "WHERE Student_Id = '" & StudentIdText.Text & "' AND Date = '" & DTPicker1.Value & "'"
            ConnDB.Execute sqlQuery     ' Execute database update
            MsgBox AttendanceUpdated, vbInformation, "Message"      ' Show update confirmation
            signalUpdateAttendance = 0      ' Reset update flag
        Else    ' Create new attendance record
            ' Build attendance insertion query
            sqlQuery = "INSERT INTO attendance_table(Student_Id, Date, Attendance) VALUES (" & _
                       "'" & StudentIdText.Text & "', '" & DTPicker1.Value & "', '" & AttendanceStatus & "')"
            ConnDB.Execute sqlQuery     ' Execute database insert
            MsgBox AttendanceMarked, vbInformation, "Message"       ' Show insert confirmation
        End If
        
        resetOption                ' Clear option button selections
        MarkAttendanceFrame.Enabled = False  ' Disable attendance frame after save
        disableOption              ' Deactivate attendance options
        visibleLabel               ' Show status labels
         
        ' Calculate overall attendance statistics from helper function
        Dim overallTotalPresent, overallTotalAbsent As Integer     ' Overall Attendance counters
        overallTotalPresent = getAttendanceCount("Present")  ' Get total present days
        overallTotalAbsent = getAttendanceCount("Absent")    ' Get total absent days
        
        ' Calculate attendance percentage
        Dim total_classes As Integer, classes_attended As Integer   ' Total classes and classes attended counts
        Dim percentage As String      ' Attendance percentage formatted as string
        Const CLASSES_PER_DAY As Integer = 6    ' Constant for classes per day
        classes_attended = overallTotalPresent * CLASSES_PER_DAY    ' Calculate total classes attended
        total_classes = (overallTotalPresent + overallTotalAbsent) * CLASSES_PER_DAY    ' Calculate total possible classes
        percentage = Round((((overallTotalPresent * CLASSES_PER_DAY) / total_classes) * 100), 2) & " %"    ' Calculate percentage
        
        ' Check if student has existing overall attendance record
        Dim RsOverallAttendanceCheck As New ADODB.Recordset     ' Overall Attendance recordset object
        sqlQuery = "SELECT COUNT(*) AS RecordExists FROM overall_attendance_table WHERE Student_Id = '" & StudentIdText.Text & "'"  ' Build record check query
        RsOverallAttendanceCheck.Open sqlQuery, ConnDB, adOpenForwardOnly, adLockReadOnly   ' Execute existence check

        If RsOverallAttendanceCheck.Fields("RecordExists").Value = 0 Then   ' Insert new overall attendance record if not exists
            ' Build overall attendance insertion query
            sqlQuery = "INSERT INTO overall_attendance_table(Student_Id, Name, Course, Class, Total_Classes, Classes_Attended, Percentage)" & _
                       "VALUES ('" & StudentIdText.Text & "', '" & NameText.Text & "', '" & CourseText.Text & "', " & _
                       "'" & ClassText.Text & "', '" & total_classes & "', '" & classes_attended & "', '" & percentage & "')"
            ConnDB.Execute sqlQuery     ' Execute database insert
        Else    ' Update existing overall attendance record
            ' Build overall attendance update query
            sqlQuery = "UPDATE overall_attendance_table SET Total_Classes = '" & total_classes & "', " & _
                       "Classes_Attended = '" & classes_attended & "', Percentage = '" & percentage & "' " & _
                       "WHERE Student_Id = '" & StudentIdText.Text & "'"
            ConnDB.Execute sqlQuery     ' Execute database update
        End If
        
        ' Cleanup overall attendance recordset
        RsOverallAttendanceCheck.Close  ' Close recordset object
        Set RsOverallAttendanceCheck = Nothing  ' Release object memory
        Exit Sub  ' Exit before error handler

' Attendance Recording Error Handler
MarkAttendanceError:
        MsgBox "Failed to mark attendance: " & Err.Description, vbCritical, "Error"  ' Display error message
        DTPicker1.Value = ""   ' Reset date selection
        resetOption            ' Clear option buttons
        MarkAttendanceFrame.Enabled = False  ' Disable attendance frame
        disableControl         ' Disable navigation controls
        clearControl           ' Clear form fields

End Sub

' --- ATTENDANCE COUNT HELPER FUNCTION ---
' Calculates attendance count for specified status
' Parameters:
'   AttendanceStatus: "Present" or "Absent"
' Returns: Count of matching records

Private Function getAttendanceCount(ByVal AttendanceStatus As String) As Integer

    Dim RsAttendanceCount As New ADODB.Recordset    ' Attendance count recordset object

    ' Build attendance count query
    sqlQuery = "SELECT COUNT(Attendance) AS overallTotalCount " & _
               "FROM attendance_table " & _
               "WHERE Attendance = '" & AttendanceStatus & "' " & _
               "AND Student_Id = '" & StudentIdText.Text & "'"
          
    ' Execute attendance count query
    RsAttendanceCount.Open sqlQuery, ConnDB, adOpenForwardOnly, adLockReadOnly
    
    ' Return count value
    getAttendanceCount = RsAttendanceCount.Fields("overallTotalCount").Value
    
    ' Cleanup recordset resources
    RsAttendanceCount.Close  ' Close recordset object
    Set RsAttendanceCount = Nothing  ' Release object memory

End Function

' --- ATTENDANCE OPTION ENABLE ---
' Activates Present/Absent option buttons

Private Sub enableOption()

    PresentOption.Enabled = True  ' Enable Present option
    AbsentOption.Enabled = True   ' Enable Absent option

End Sub

' --- ATTENDANCE OPTION DISABLE ---
' Deactivates Present/Absent option buttons

Private Sub disableOption()

    PresentOption.Enabled = False  ' Disable Present option
    AbsentOption.Enabled = False   ' Disable Absent option

End Sub

' --- ATTENDANCE OPTION RESET ---
' Clears selection from option buttons

Private Sub resetOption()

    PresentOption.Value = False  ' Deselect Present option
    AbsentOption.Value = False   ' Deselect Absent option

End Sub

' --- STATUS LABEL VISIBLE ---
' Shows attendance status labels

Private Sub visibleLabel()

    UpdateAttendanceLabel.Visible = True      ' Show update label
    MarkedAttendanceLabel.Visible = True      ' Show marked status label

End Sub

' --- STATUS LABEL HIDE ---
' Hides attendance status labels

Private Sub hideLabel()

    UpdateAttendanceLabel.Visible = False     ' Hide update label
    MarkedAttendanceLabel.Visible = False     ' Hide marked status label

End Sub

' --- NAVIGATION CONTROL ENABLE ---
' Activates record navigation buttons

Private Sub enableControl()

    FirstButton.Enabled = True    ' Enable First button
    NextButton.Enabled = True     ' Enable Next button
    PreviousButton.Enabled = True ' Enable Previous button
    LastButton.Enabled = True     ' Enable Last button

End Sub

' --- NAVIGATION CONTROL DISABLE ---
' Deactivates record navigation buttons

Private Sub disableControl()

    FirstButton.Enabled = False    ' Disable First button
    NextButton.Enabled = False     ' Disable Next button
    PreviousButton.Enabled = False ' Disable Previous button
    LastButton.Enabled = False     ' Disable Last button

End Sub

' --- UI CONTROL CLEAR ---
' Resets student data display

Private Sub clearControl()

    StudentIdText.Text = ""          ' Clear student id field
    NameText.Text = ""               ' Clear student name field
    StudentImage.Picture = LoadPicture("")  ' Clear student photo
    CourseText.Text = ""             ' Clear course field
    ClassText.Text = ""              ' Clear class field
    hideLabel                        ' Hide status labels

End Sub

' --- STUDENT RECORDSET CLEANUP ---
' Safely closes student details recordset

Private Sub closeRsAttendanceDetails()

    ' Close recordset if active
    If RsStudentDetails.State = adStateOpen Then
        RsStudentDetails.Close     ' Release database resources
        Set RsStudentDetails = Nothing  ' Free object memory
    End If

End Sub

' --- ATTENDANCE RECORDSET CLEANUP ---
' Safely closes attendance status recordset

Private Sub closeRsAttendanceStatus()

    ' Close recordset if active
    If RsAttendanceStatus.State = adStateOpen Then
        RsAttendanceStatus.Close     ' Release database resources
        Set RsAttendanceStatus = Nothing  ' Free object memory
    End If

End Sub

' --- FORM UNLOAD EVENT ---
' Handles cleanup when form is closed

Private Sub Form_Unload(Cancel As Integer)

    isOpenMarkAttendance = False  ' Reset mark attendance form state flag
    
    closeRsAttendanceDetails   ' Cleanup student recordset
    closeRsAttendanceStatus    ' Cleanup attendance recordset

    ' Close database connection if active
    If ConnDB.State = adStateOpen Then
        ConnDB.Close    ' Release database connection
        Set ConnDB = Nothing    ' Free object memory
    End If

End Sub
