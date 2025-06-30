VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddStudent 
   BackColor       =   &H00400040&
   Caption         =   "Add Student"
   ClientHeight    =   9735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16770
   Icon            =   "frmAddStudent.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   9735
   ScaleWidth      =   16770
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton SaveStudentButton 
      BackColor       =   &H0000FF00&
      Caption         =   "SAVE"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton AddStudentButton 
      BackColor       =   &H000000FF&
      Caption         =   "ADD"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9120
      Width           =   2535
   End
   Begin VB.ComboBox GenderCombo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      ItemData        =   "frmAddStudent.frx":5A5A2
      Left            =   14160
      List            =   "frmAddStudent.frx":5A5AF
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2520
      Width           =   3495
   End
   Begin VB.ComboBox ClassCombo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      ItemData        =   "frmAddStudent.frx":5A5CE
      Left            =   14160
      List            =   "frmAddStudent.frx":5A5E4
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4920
      Width           =   3495
   End
   Begin VB.ComboBox CourseCombo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      ItemData        =   "frmAddStudent.frx":5A61E
      Left            =   14160
      List            =   "frmAddStudent.frx":5A62B
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3720
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox AddressText 
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
      Left            =   14160
      TabIndex        =   8
      Top             =   7320
      Width           =   3495
   End
   Begin VB.CommandButton UploadPhotoButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "Upload"
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
      TabIndex        =   3
      Top             =   7320
      Width           =   2295
   End
   Begin VB.TextBox MobileNoText 
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
      Left            =   14160
      MaxLength       =   10
      TabIndex        =   7
      Top             =   6120
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
      Left            =   6960
      TabIndex        =   2
      Top             =   3600
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
      Left            =   6960
      TabIndex        =   1
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label Label8 
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
      Left            =   10560
      TabIndex        =   16
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Image StudentImage 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   2415
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label9 
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
      Left            =   10560
      TabIndex        =   15
      Top             =   7440
      Width           =   2895
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
      Left            =   3240
      TabIndex        =   14
      Top             =   5160
      Width           =   2895
   End
   Begin VB.Label Label7 
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
      Left            =   10560
      TabIndex        =   13
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label6 
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
      Left            =   10560
      TabIndex        =   12
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Label Label4 
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
      Left            =   10560
      TabIndex        =   11
      Top             =   3840
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
      Left            =   3240
      TabIndex        =   10
      Top             =   3720
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
      Left            =   3240
      TabIndex        =   9
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Student"
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
      Left            =   9840
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "frmAddStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================== ADD STUDENT FORM =====================
' Interface for adding new student records to the system
' Key Features:
'   - Student photo management
'   - Comprehensive input validation
'   - Gender/course/class selection
'   - Database integration with error handling
' ===========================================================

Option Explicit

' --- GLOBAL VARIABLES ---

Dim ConnDB As New ADODB.Connection      ' Database connection object
Dim imagePath As String                 ' Student photo storage path
Public isOpenAddStudent As Boolean      ' Add Student form open state flag

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
        Unload Me   ' Terminate add student form
    
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

' --- ADD STUDENT BUTTON CLICK ---
' Prepares form for new student entry

Private Sub AddStudentButton_Click()

    enableControl     ' Activate student data input controls

    StudentIdText.Text = ""      ' Reset student id field
    NameText.Text = ""           ' Reset name field
    StudentImage.Picture = LoadPicture("")  ' Reset student image field
    GenderCombo.ListIndex = -1   ' Reset gender selection field
    CourseCombo.ListIndex = -1   ' Reset course selection field
    ClassCombo.ListIndex = -1    ' Reset class selection field
    MobileNoText.Text = ""       ' Reset mobile number field
    AddressText.Text = ""        ' Reset address field
    
    StudentIdText.SetFocus       ' Focus student id field for input
    SaveStudentButton.Enabled = True   ' Enable save functionality
    AddStudentButton.Enabled = False   ' Disable add button during operation

End Sub

' --- SAVE STUDENT BUTTON CLICK ---
' Validates and saves new student record

Private Sub SaveStudentButton_Click()

    ' Check for empty required fields
    If Trim(StudentIdText.Text) = "" Or Trim(NameText.Text) = "" Or Trim(imagePath) = "" Or _
       Trim(GenderCombo.Text) = "" Or Trim(CourseCombo.Text) = "" Or Trim(ClassCombo.Text) = "" Or _
       Trim(AddressText.Text) = "" Or Trim(MobileNoText.Text) = "" Then
            MsgBox "All Fields are Mandatory", vbExclamation, "Warning"  ' Show warning
            Exit Sub            ' Abort save operation
    End If
    
    ' Verify 10-digit mobile number format
    If Len(MobileNoText.Text) <> 10 Then
        MsgBox "Please enter a valid 10-digit mobile number.", vbExclamation, "Warning"  ' Show warning
        MobileNoText.SetFocus  ' Focus mobile field for correction
        Exit Sub               ' Abort save operation
    End If
    
    Dim sqlQuery As String  ' SQL command builder

    ' Build student insertion query
    sqlQuery = "INSERT INTO student_table (Student_Id, Name, Photo, Gender, Course, Class, Mobile_No, Address) VALUES (" & _
               "'" & StudentIdText.Text & "', '" & Trim(NameText.Text) & "', '" & imagePath & "', '" & GenderCombo.Text & "', " & _
               "'" & CourseCombo.Text & "', '" & ClassCombo.Text & "', '" & MobileNoText.Text & "', '" & Trim(AddressText.Text) & "')"

    On Error GoTo SaveError  ' Redirect runtime errors to SaveError handler
    
        ConnDB.Execute sqlQuery  ' Execute database command
        MsgBox "New Student Added", vbInformation, "Message"  ' Show success notification
        disableControl           ' Deactivate student data input controls
        AddStudentButton.Enabled = True    ' Reactivate add button
        SaveStudentButton.Enabled = False  ' Deactivate save button
        Exit Sub                 ' Exit before error handler
        
' Save Error Handler
SaveError:
        disableControl          ' Deactivate student data input controls
        AddStudentButton.Enabled = True    ' Reactivate add button
        SaveStudentButton.Enabled = False  ' Deactivate save button
        MsgBox "Failed to add student: " & Err.Description, vbCritical, "Error"  ' Show error details

End Sub

' --- CONTROL ENABLE ROUTINE ---
' Activates form input controls

Private Sub enableControl()

    StudentIdText.Enabled = True       ' Activate student id field
    NameText.Enabled = True            ' Activate name field
    MobileNoText.Enabled = True        ' Activate mobile number field
    AddressText.Enabled = True         ' Activate address field
    GenderCombo.Enabled = True         ' Activate gender selection
    CourseCombo.Enabled = True         ' Activate course selection
    ClassCombo.Enabled = True          ' Activate class selection
    StudentImage.Enabled = True        ' Activate image display
    UploadPhotoButton.Enabled = True   ' Activate photo upload button

End Sub

' --- CONTROL DISABLE ROUTINE ---
' Deactivates form input controls

Private Sub disableControl()

    StudentIdText.Enabled = False       ' Deactivate student id field
    NameText.Enabled = False            ' Deactivate name field
    MobileNoText.Enabled = False        ' Deactivate mobile number field
    AddressText.Enabled = False         ' Deactivate address field
    GenderCombo.Enabled = False         ' Deactivate gender selection
    CourseCombo.Enabled = False         ' Deactivate course selection
    ClassCombo.Enabled = False          ' Deactivate class selection
    StudentImage.Enabled = False        ' Deactivate image display
    UploadPhotoButton.Enabled = False   ' Deactivate photo upload button

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

' --- FORM UNLOAD EVENT ---
' Handles cleanup when form is closed

Private Sub Form_Unload(Cancel As Integer)
    
    isOpenAddStudent = False  ' Reset add student form open state flag
    
    ' Close database connection if open
    If ConnDB.State = adStateOpen Then
        ConnDB.Close    ' Close active connection
        Set ConnDB = Nothing    ' Release object memory
    End If

End Sub
