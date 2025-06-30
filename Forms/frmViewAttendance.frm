VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmViewAttendance 
   BackColor       =   &H00404000&
   Caption         =   "View Attendance"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17130
   Icon            =   "frmViewAttendance.frx":0000
   LinkTopic       =   "Form12"
   ScaleHeight     =   9630
   ScaleWidth      =   17130
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CloseButton 
      BackColor       =   &H000000FF&
      Caption         =   "CLOSE"
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
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton ShowAttendanceButton 
      BackColor       =   &H0000FF00&
      Caption         =   "SHOW"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   10320
      TabIndex        =   1
      Top             =   4080
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
      Left            =   10320
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label0 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "View Attendance"
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
      Left            =   7440
      TabIndex        =   0
      Top             =   600
      Width           =   7575
   End
End
Attribute VB_Name = "frmViewAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================== VIEW ATTENDANCE FORM =====================
' Interface for viewing and generating attendance reports
' Key Features:
'   - Date-based attendance filtering
'   - Dynamic report generation
'   - Attendance statistics calculation
'   - Database resource management
' ===============================================================

Option Explicit

' --- GLOBAL VARIABLES ---

Dim ConnDB As New ADODB.Connection          ' Database connection object
Dim RsDataReport As New ADODB.Recordset     ' Report data recordset object
Public isOpenViewAttendance As Boolean      ' View Attendance form open state flag

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
        Unload Me   ' Terminate view attendance form
    
End Sub

' --- DATE PICKER CHANGE EVENT ---
' Enables report generation when date is selected

Private Sub DTPicker1_Change()

    ShowAttendanceButton.Enabled = True  ' Enable show attendance button

End Sub

' --- SHOW ATTENDANCE BUTTON CLICK ---
' Generates attendance report for selected date

Private Sub ShowAttendanceButton_Click()

    ' Close attendance report if already open
    If isOpenAttendanceDetails Then
        Unload dRepAttendanceDetails  ' Terminate attendance report
    End If

    ' Validate date selection
    If DTPicker1.Value = "" Then
        MsgBox "Select Date", vbExclamation, "Warning"  ' Show warning
        Exit Sub            ' Abort report generation
    End If
    
    Dim totalPresent, totalAbsent As Integer  ' Attendance counters
    Dim sqlQuery As String                    ' SQL query builder
    
    ' Get attendance statistics from helper function
    totalPresent = getAttendanceCount("Present")  ' Count present students
    totalAbsent = getAttendanceCount("Absent")    ' Count absent students
    
    closeRsDataReport  ' Cleanup previous recordsets
    
    ' Build attendance details query
    sqlQuery = "SELECT student_table.Student_Id, student_table.Name, student_table.Course, " & _
               "student_table.Class, attendance_table.Attendance " & _
               "FROM student_table, attendance_table " & _
               "WHERE attendance_table.Date = '" & DTPicker1.Value & "' " & _
               "AND student_table.Student_Id = attendance_table.Student_Id " & _
               "ORDER BY student_table.Student_Id"
              
    ' Execute attendance details query
    RsDataReport.Open sqlQuery, ConnDB, adOpenStatic, adLockReadOnly

    ' Handle empty result set
    If RsDataReport.RecordCount = 0 Then
        MsgBox "No Attendance was taken on " & DTPicker1.Value, vbExclamation, "Warning"  ' Show warning
        Exit Sub            ' Abort report generation
    End If
    
    ' Configure attendance details report
    Set dRepAttendanceDetails.DataSource = RsDataReport     ' Bind data to report
    dRepAttendanceDetails.Sections("Section4").Controls.Item("DateLabel").Caption = "" & DTPicker1.Value    ' Set report date
    dRepAttendanceDetails.Sections("Section5").Controls.Item("PresentLabel").Caption = "" & totalPresent    ' Set present count
    dRepAttendanceDetails.Sections("Section5").Controls.Item("AbsentLabel").Caption = "" & totalAbsent      ' Set absent count
    dRepAttendanceDetails.Show       ' Display attendance detailsreport
    
    isOpenAttendanceDetails = True   ' Set attendance details report open state flag
    ShowAttendanceButton.Enabled = False  ' Disable show attendance button

End Sub

' --- CLOSE BUTTON CLICK ---
' Terminates View Attendance form

Private Sub CloseButton_Click()

    Unload Me  ' Terminate view attendance interface

End Sub

' --- ATTENDANCE COUNT HELPER FUNCTION ---
' Calculates attendance count for specified status
' Parameters:
'   AttendanceStatus: "Present" or "Absent"
' Returns: Count of matching records

Private Function getAttendanceCount(ByVal AttendanceStatus As String) As Integer

    closeRsDataReport  ' Cleanup previous recordsets
        
    Dim sqlQuery As String  ' SQL query builder
    
    ' Build attendance count query
    sqlQuery = "SELECT COUNT(Attendance) AS totalCount " & _
               "FROM attendance_table " & _
               "WHERE Attendance = '" & AttendanceStatus & "' " & _
               "AND Date = '" & DTPicker1.Value & "' "
          
    ' Execute attendance count query
    RsDataReport.Open sqlQuery, ConnDB, adOpenForwardOnly, adLockReadOnly
    
    ' Return attendance count value
    getAttendanceCount = RsDataReport.Fields("totalCount").Value

End Function

' --- RECORDSET CLEANUP ---
' Safely closes data report recordset

Private Sub closeRsDataReport()
    
    ' Close database recordset if open
    If RsDataReport.State = adStateOpen Then
        RsDataReport.Close     ' Close active recordset
        Set RsDataReport = Nothing    ' Release object memory
    End If
    
End Sub

' --- FORM UNLOAD EVENT ---
' Handles cleanup when form is closed

Private Sub Form_Unload(Cancel As Integer)

    isOpenViewAttendance = False  ' Reset view attendance open state flag
    
    closeRsDataReport  ' Execute recordset cleanup
    
    ' Close database connection if open
    If ConnDB.State = adStateOpen Then
        ConnDB.Close    ' Close active connection
        Set ConnDB = Nothing    ' Release object memory
    End If
    
    ' Close attendance report if open
    If isOpenAttendanceDetails Then
        Unload dRepAttendanceDetails  ' Terminate attendance details report
    End If

End Sub
