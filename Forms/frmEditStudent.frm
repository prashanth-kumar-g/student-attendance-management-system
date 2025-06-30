VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditStudent 
   BackColor       =   &H00400040&
   Caption         =   "Edit Student"
   ClientHeight    =   9675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16575
   Icon            =   "frmEditStudent.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   9675
   ScaleWidth      =   16575
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox GenderCombo 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      ItemData        =   "frmEditStudent.frx":5A5A2
      Left            =   13920
      List            =   "frmEditStudent.frx":5A5AF
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   3480
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox ClassCombo 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      ItemData        =   "frmEditStudent.frx":5A5CE
      Left            =   13920
      List            =   "frmEditStudent.frx":5A5E4
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   5880
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox CourseCombo 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      ItemData        =   "frmEditStudent.frx":5A61E
      Left            =   13920
      List            =   "frmEditStudent.frx":5A62B
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   4680
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton DeleteStudentButton 
      BackColor       =   &H000000FF&
      Caption         =   "DELETE"
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
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   10080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton ProceedSearchButton 
      BackColor       =   &H0000FF00&
      Caption         =   "PROCEED"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox SearchStudentText 
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
      Left            =   4440
      TabIndex        =   21
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox StudentIdText 
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
      Left            =   6360
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox NameText 
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
      Left            =   6360
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox MobileNoText 
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
      Left            =   13920
      MaxLength       =   10
      TabIndex        =   4
      Top             =   7080
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton UpdateStudentButton 
      BackColor       =   &H0000FF00&
      Caption         =   "UPDATE"
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
      Height          =   855
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton UploadPhotoButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "Upload"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox AddressText 
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
      Left            =   13920
      TabIndex        =   1
      Top             =   8280
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label EditLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   17640
      TabIndex        =   29
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label AttributeLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Index           =   4
      Left            =   10440
      TabIndex        =   28
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the Student ID To Edit"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   3600
      TabIndex        =   23
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label EditLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   10080
      TabIndex        =   20
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label EditLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   10080
      TabIndex        =   19
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label EditLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   8880
      TabIndex        =   18
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label EditLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   17640
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label EditLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   17640
      TabIndex        =   16
      Top             =   7320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label EditLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   17640
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label EditLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   17640
      TabIndex        =   14
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label AttributeLabel 
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
      Index           =   1
      Left            =   2760
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label AttributeLabel 
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
      Index           =   2
      Left            =   2760
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label AttributeLabel 
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
      Index           =   5
      Left            =   10440
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label AttributeLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No."
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
      Index           =   7
      Left            =   10440
      TabIndex        =   10
      Top             =   7200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label AttributeLabel 
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
      Index           =   6
      Left            =   10440
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label AttributeLabel 
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
      Index           =   3
      Left            =   2760
      TabIndex        =   8
      Top             =   6240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label AttributeLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Index           =   8
      Left            =   10440
      TabIndex        =   7
      Top             =   8400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image StudentImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Student"
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
      Height          =   855
      Left            =   10920
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "frmEditStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================== EDIT STUDENT FORM =====================
' Interface for modifying and deleting student records
' Key Features:
'   - Student record search and retrieval
'   - Selective field editing with visual indicators
'   - Comprehensive input validation
'   - Update/delete operations with confirmation
' ============================================================

Option Explicit

' --- GLOBAL VARIABLES ---

Dim ConnDB As New ADODB.Connection          ' Database connection object
Dim RsStudentDetails As New ADODB.Recordset ' Student details recordset object
Dim imagePath As String                     ' Student photo storage path
Dim msg As VbMsgBoxResult                   ' Message box response storage
Dim sqlQuery As String                      ' SQL command builder
Dim i As Integer                            ' Loop control variable
Public isOpenEditStudent As Boolean         ' Edit Student form open state flag

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
        hideControl   ' Initialize UI with student details hidden
        Exit Sub      ' Exit before error handler
    
' Database Connection Error Handler
ConnectionError:
        MsgBox "Database connection failed: " & Err.Description, vbCritical, "Error"  ' Show error details
        Unload Me   ' Terminate edit student form

End Sub

' --- PROCEED SEARCH BUTTON CLICK ---
' Retrieves student details based on search input

Private Sub ProceedSearchButton_Click()

    ' Check for empty student ids
    If Trim(SearchStudentText.Text) = "" Then
        clearControl     ' Reset input fields
        MsgBox "Enter Student Id", vbExclamation, "Warning"  ' Show warning
        hideControl      ' Hide student details section
        Exit Sub         ' Abort search operation
    End If
    
    ' Query database for student record
    RsStudentDetails.Open "select * from student_table where Student_Id='" & SearchStudentText & "'", ConnDB, adOpenStatic, adLockReadOnly
    
    ' Check if record exists
    If Not RsStudentDetails.EOF = True Then
        ' Populate form fields with student data
        StudentIdText.Text = RsStudentDetails.Fields("Student_Id").Value    ' Set student id field
        NameText.Text = RsStudentDetails.Fields("Name").Value               ' Set name field
        imagePath = RsStudentDetails.Fields("Photo").Value                  ' Store photo path
        StudentImage.Picture = LoadPicture(App.Path & "\" & RsStudentDetails.Fields("Photo").Value)  ' Set photo field
        GenderCombo.Text = RsStudentDetails.Fields("Gender").Value          ' Set gender field
        CourseCombo.Text = RsStudentDetails.Fields("Course").Value          ' Set course field
        ClassCombo.Text = RsStudentDetails.Fields("Class").Value            ' Set class field
        MobileNoText.Text = RsStudentDetails.Fields("Mobile_No").Value      ' Set mobile number field
        AddressText.Text = RsStudentDetails.Fields("Address").Value         ' Set address field
        
        showControl      ' Display student details section
        disableControl   ' Deactivate input fields initially
        UpdateStudentButton.Enabled = False  ' Disable update button
        DeleteStudentButton.Enabled = True   ' Enable delete button
        MsgBox "Student Found", vbInformation, "Message"  ' Show success notification
    Else
        clearControl     ' Reset input fields
        MsgBox "Invalid Student Id", vbCritical, "Error"  ' Show error
        hideControl      ' Hide student details section
    End If
    
    closeRsStudentDetails  ' Cleanup recordset resources

End Sub

' --- EDIT LABEL CLICK ---
' Enables editing for specific student attribute

Private Sub EditLabel_Click(index As Integer)
    
    UpdateStudentButton.Enabled = True  ' Enable update functionality
    EditLabel(index).Forecolor = &HFF&  ' Change label color to red (visual edit indicator)
    
    ' Activate specific field based on index
    If index = 1 Then
        StudentIdText.Enabled = True   ' Activate student id field
        StudentIdText.SetFocus         ' Focus student id field for editing
    ElseIf index = 2 Then
        NameText.Enabled = True        ' Activate name field
        NameText.SetFocus              ' Focus name field for editing
    ElseIf index = 3 Then
        UploadPhotoButton.Enabled = True  ' Activate photo upload button
        UploadPhotoButton.SetFocus        ' Focus upload button for operation
    ElseIf index = 4 Then
        GenderCombo.Enabled = True     ' Activate gender selection
        GenderCombo.SetFocus           ' Focus gender combo for editing
    ElseIf index = 5 Then
        CourseCombo.Enabled = True     ' Activate course selection
        CourseCombo.SetFocus           ' Focus course combo for editing
    ElseIf index = 6 Then
        ClassCombo.Enabled = True      ' Activate class selection
        ClassCombo.SetFocus            ' Focus class combo for editing
    ElseIf index = 7 Then
        MobileNoText.Enabled = True    ' Activate mobile number field
        MobileNoText.SetFocus          ' Focus mobile field for editing
    ElseIf index = 8 Then
        AddressText.Enabled = True     ' Activate address field
        AddressText.SetFocus           ' Focus address field for editing
    End If

End Sub

' --- UPLOAD PHOTO BUTTON CLICK ---
' Handles student photo selection and storage

Private Sub UploadPhotoButton_Click()

    Dim sourcePath As String  ' Original image location
    Dim destPath As String    ' Application storage path

    CommonDialog1.Filter = "Image Files (*.jpg;*.jpeg)|*.jpg;*.jpeg"    ' Configure image file filter for dialog
    CommonDialog1.ShowOpen    ' Display file selection dialog

    ' Process selected file
    If CommonDialog1.FileName <> "" Then
        sourcePath = CommonDialog1.FileName  ' Capture source path
        destPath = App.Path & "\Images\" & Dir(sourcePath)  ' Build destination path
        
        ' Avoid self-copy operation
        If LCase(sourcePath) <> LCase(destPath) Then
            If Dir(destPath) <> "" Then Kill destPath   ' Delete existing duplicate in application path
            FileCopy sourcePath, destPath  ' Copy file to application directory if not exists duplicate
        End If

        imagePath = "Images\" & Dir(sourcePath)  ' Set relative image path
        StudentImage.Picture = LoadPicture(App.Path & "\" & imagePath)  ' Display uploaded image
    End If

End Sub

' --- UPDATE STUDENT BUTTON CLICK ---
' Validates and saves modified student record

Private Sub UpdateStudentButton_Click()

    ' Check for empty required fields
    If Trim(StudentIdText.Text) = "" Or Trim(NameText.Text) = "" Or Trim(MobileNoText.Text) = "" Or Trim(AddressText.Text) = "" Or _
       Trim(GenderCombo.Text) = "" Or Trim(CourseCombo.Text) = "" Or Trim(ClassCombo.Text) = "" Or Trim(imagePath) = "" Then
            MsgBox "All Fields are Mandatory", vbExclamation, "Warning"  ' Show warning
            Exit Sub            ' Abort update operation
    End If
    
    ' Verify 10-digit mobile number format
    If Len(MobileNoText.Text) <> 10 Then
        MsgBox "Please enter a valid 10-digit mobile number.", vbExclamation, "Warning"  ' Show warning
        MobileNoText.SetFocus  ' Focus mobile field for correction
        Exit Sub               ' Abort update operation
    End If
    
    ' Confirm update operation
    msg = MsgBox("Are you sure want to Update", vbQuestion + vbYesNo, "Warning")  ' Confirmation dialog (yes/no)
        
    ' If selected yes proceed to update
    If msg = vbYes Then
        ' Build student update query
        sqlQuery = "UPDATE student_table SET Student_Id='" & StudentIdText.Text & "', Name='" & Trim(NameText.Text) & "', " & _
                   "Photo='" & imagePath & "', Gender='" & GenderCombo.Text & "',Course='" & CourseCombo.Text & "', " & _
                   "Class='" & ClassCombo.Text & "', Mobile_No='" & MobileNoText.Text & "', Address='" & Trim(AddressText.Text) & "' " & _
                   "WHERE Student_Id='" & SearchStudentText.Text & "' "
      
        On Error GoTo UpdateError  ' Redirect runtime errors to UpdateError handler
        
            ConnDB.Execute sqlQuery  ' Execute update command
            MsgBox "Student Information Updated", vbInformation, "Message"  ' Show success notification
            disableControl           ' Deactivate input fields
            UpdateStudentButton.Enabled = False  ' Disable update button
            Exit Sub                 ' Exit before error handler
            
' Update Error Handler
UpdateError:
        MsgBox "Failed to update student: " & Err.Description, vbCritical, "Error"  ' Show error details
        disableControl          ' Deactivate input fields
        UpdateStudentButton.Enabled = False  ' Disable update button
    End If

End Sub

' --- DELETE STUDENT BUTTON CLICK ---
' Removes student record from system

Private Sub DeleteStudentButton_Click()

    ' Verify student id exists
    If SearchStudentText.Text = "" And StudentIdText.Text = "" Then
        MsgBox "Student Id is Necessary for deletion", vbExclamation, "Warning"  ' Show warning
        hideControl   ' Hide student details section
        Exit Sub      ' Abort delete operation
    End If
 
    ' Confirm delete operation
    msg = MsgBox("Are you sure want to Delete", vbExclamation + vbYesNo, "Warning")  ' Confirmation dialog (yes/no)
    
    ' If selected yes proceed to delete
    If msg = vbYes Then
    ' Build student deletion query
        sqlQuery = "DELETE FROM student_table WHERE Student_Id='" & SearchStudentText.Text & "'"
        
        On Error GoTo DeleteError  ' Redirect runtime errors to DeleteError handler
        
            ConnDB.Execute sqlQuery  ' Execute delete command
            MsgBox "Student Deleted from Records", vbInformation, "Message"  ' Show success notification
            clearControl             ' Reset input fields
            DeleteStudentButton.Enabled = False  ' Disable delete button
            UpdateStudentButton.Enabled = False  ' Disable update button
            Exit Sub                 ' Exit before error handler
            
' Delete Error Handler
DeleteError:
        MsgBox "Failed to delete student: " & Err.Description, vbCritical, "Error"  ' Show error details
        clearControl            ' Reset input fields
        DeleteStudentButton.Enabled = False  ' Disable delete button
    End If

End Sub

' --- SHOW CONTROL ROUTINE ---
' Displays student details section

Private Sub showControl()

    ' Make attribute labels and edit controls visible
    For i = 1 To 8
        AttributeLabel(i).Visible = True  ' Show field labels
        EditLabel(i).Visible = True       ' Show edit triggers
    Next
    
    ' Make input fields and buttons visible
    StudentIdText.Visible = True         ' Show student id field
    NameText.Visible = True              ' Show name field
    MobileNoText.Visible = True          ' Show mobile field
    AddressText.Visible = True           ' Show address field
    GenderCombo.Visible = True           ' Show gender selection field
    CourseCombo.Visible = True           ' Show course selection field
    ClassCombo.Visible = True            ' Show class selection field
    StudentImage.Visible = True          ' Show photo display
    UploadPhotoButton.Visible = True     ' Show upload button
    UpdateStudentButton.Visible = True   ' Show update button
    DeleteStudentButton.Visible = True   ' Show delete button

End Sub

' --- HIDE CONTROL ROUTINE ---
' Conceals student details section

Private Sub hideControl()
    
    ' Hide attribute labels and edit controls
    For i = 1 To 8
        AttributeLabel(i).Visible = False  ' Hide field labels
        EditLabel(i).Visible = False       ' Hide edit triggers
    Next
    
    ' Hide input fields and buttons
    StudentIdText.Visible = False         ' Hide student id field
    NameText.Visible = False              ' Hide name field
    MobileNoText.Visible = False          ' Hide mobile field
    AddressText.Visible = False           ' Hide address field
    GenderCombo.Visible = False           ' Hide gender selection field
    CourseCombo.Visible = False           ' Hide course selection field
    ClassCombo.Visible = False            ' Hide class selection field
    StudentImage.Visible = False          ' Hide photo display
    UploadPhotoButton.Visible = False     ' Hide upload button
    UpdateStudentButton.Visible = False   ' Hide update button
    DeleteStudentButton.Visible = False   ' Hide delete button

End Sub

' --- DISABLE CONTROL ROUTINE ---
' Deactivates input fields and resets edit indicators

Private Sub disableControl()

    ' Reset edit label colors to white
    For i = 1 To 8
        EditLabel(i).Forecolor = &HFFFFFF  ' Reset visual indicator
    Next
    
    ' Deactivate input fields
    StudentIdText.Enabled = False        ' Lock student id field
    NameText.Enabled = False             ' Lock name field
    MobileNoText.Enabled = False         ' Lock mobile field
    AddressText.Enabled = False          ' Lock address field
    GenderCombo.Enabled = False          ' Lock gender selection field
    CourseCombo.Enabled = False          ' Lock course selection field
    ClassCombo.Enabled = False           ' Lock class selection field
    UploadPhotoButton.Enabled = False    ' Lock upload button

End Sub

' --- CLEAR CONTROL ROUTINE ---
' Resets input fields to empty state

Private Sub clearControl()

    ' Clear all input fields
    StudentIdText.Text = ""      ' Reset student id field
    NameText.Text = ""           ' Reset name field
    MobileNoText.Text = ""       ' Reset mobile field
    AddressText.Text = ""        ' Reset address field
    GenderCombo.ListIndex = -1   ' Reset gender selection field
    CourseCombo.ListIndex = -1   ' Reset course selection field
    ClassCombo.ListIndex = -1    ' Reset class selection field
    StudentImage.Picture = LoadPicture("")  ' Reset photo field

End Sub

' --- SEARCH STUDENT TEXT KEYPRESS VALIDATION ---
' Restricts input to alphanumeric characters

Private Sub SearchStudentText_KeyPress(KeyAscii As Integer)

    ' Allow: A-Z, a-z, 0-9, backspace
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        ' Valid input - allow processing
    Else
        KeyAscii = 0  ' Block invalid characters
    End If

End Sub

' --- STUDENT ID TEXT KEYPRESS VALIDATION ---
' Restricts input to alphanumeric characters

Private Sub StudentIdText_KeyPress(KeyAscii As Integer)

    ' Allow: A-Z, a-z, 0-9, backspace
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        ' Valid input - allow processing
    Else
        KeyAscii = 0  ' Block invalid characters
    End If

End Sub

' --- NAME TEXT KEYPRESS VALIDATION ---
' Restricts input to alphabetic characters and spaces

Private Sub NameText_KeyPress(KeyAscii As Integer)
    
    ' Allow: A-Z, a-z, space, backspace
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
        ' Valid input - allow processing
    Else
        KeyAscii = 0  ' Block invalid characters
    End If

End Sub

' --- MOBILE NUMBER TEXT KEYPRESS VALIDATION ---
' Restricts input to numeric digits

Private Sub MobileNoText_KeyPress(KeyAscii As Integer)

    ' Allow: 0-9, backspace
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        ' Valid input - allow processing
    Else
        KeyAscii = 0  ' Block invalid characters
    End If

End Sub

' --- RECORDSET CLEANUP ---
' Safely closes student details recordset

Private Sub closeRsStudentDetails()
    
    ' Close recordset if active
    If RsStudentDetails.State = adStateOpen Then
        RsStudentDetails.Close     ' Release database resources
        Set RsStudentDetails = Nothing  ' Free object memory
    End If
    
End Sub

' --- FORM UNLOAD EVENT ---
' Performs cleanup when form closes

Private Sub Form_Unload(Cancel As Integer)

    isOpenEditStudent = False  ' Reset edit student form open state flag
    
    ' Close database connection if active
    If ConnDB.State = adStateOpen Then
        ConnDB.Close    ' Release database connection
        Set ConnDB = Nothing    ' Free object memory
    End If

End Sub
