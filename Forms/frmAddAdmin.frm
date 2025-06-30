VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddAdmin 
   BackColor       =   &H00400040&
   Caption         =   "Add Admin"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16770
   Icon            =   "frmAddAdmin.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   9645
   ScaleWidth      =   16770
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton AddAdminButton 
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8160
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13320
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton SaveAdminButton 
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
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8160
      Width           =   2535
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
      Left            =   13920
      TabIndex        =   5
      Top             =   2520
      Width           =   3495
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
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   4
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
      TabIndex        =   1
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox AdminIdText 
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
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox PasswordText 
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
      IMEMode         =   3  'DISABLE
      Left            =   6960
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Image AdminImage 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   2295
      Left            =   13920
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Photo"
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
      Left            =   10320
      TabIndex        =   14
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label6 
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
      Left            =   10320
      TabIndex        =   13
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label5 
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
      Left            =   3360
      TabIndex        =   12
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Label Label2 
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
      Left            =   3360
      TabIndex        =   11
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Admin ID"
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
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   3360
      TabIndex        =   9
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Admin"
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
      Left            =   9600
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
End
Attribute VB_Name = "frmAddAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================== ADD ADMIN FORM =====================
' Interface for adding new administrator accounts to the system
' Key Features:
'   - Photo upload and management
'   - Comprehensive input validation
'   - Secure password enforcement
'   - Database integration with error handling
' =========================================================

Option Explicit

' --- GLOBAL VARIABLES ---

Dim ConnDB As New ADODB.Connection      ' Database connection object
Dim imagePath As String                 ' Administrator photo storage path
Public isOpenAddAdmin As Boolean        ' Add Admin form open state flag

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
        Unload Me   ' Terminate add admin form
    
End Sub

' --- UPLOAD PHOTO BUTTON CLICK ---
' Handles administrator photo selection and storage

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
        AdminImage.Picture = LoadPicture(App.Path & "\" & imagePath)  ' Display uploaded image
    End If

End Sub

' --- ADD ADMIN BUTTON CLICK ---
' Prepares form for new admin entry

Private Sub AddAdminButton_Click()

    enableControl     ' Activate admin data input controls
    
    NameText.Text = ""          ' Reset name field
    AdminIdText.Text = ""       ' Reset admin id field
    PasswordText.Text = ""      ' Reset password field
    MobileNoText.Text = ""      ' Reset mobile number field
    AddressText.Text = ""       ' Reset address field
    AdminImage.Picture = LoadPicture("")  ' Reset admin image field
    
    NameText.SetFocus           ' Focus name field for input
    SaveAdminButton.Enabled = True   ' Enable save functionality
    AddAdminButton.Enabled = False   ' Disable add button during operation

End Sub

' --- SAVE ADMIN BUTTON CLICK ---
' Validates and saves new administrator record

Private Sub SaveAdminButton_Click()

    NameText.SetFocus   ' Ensure focus starts at name field

    ' Check for empty required fields
    If Trim(NameText.Text) = "" Or Trim(AdminIdText.Text) = "" Or Trim(PasswordText.Text) = "" Or _
       Trim(MobileNoText.Text) = "" Or Trim(AddressText.Text) = "" Or Trim(imagePath) = "" Then
            MsgBox "All Fields are Mandatory", vbExclamation, "Warning"  ' Show warning
            Exit Sub            ' Abort save operation
    End If
    
    ' Enforce secure password policy
    If Not validatePassword(PasswordText.Text) Then
        MsgBox "Password must be at least 8 characters and include:" & vbCrLf & _
               "- at least 1 uppercase letter" & vbCrLf & _
               "- at least 1 lowercase letter" & vbCrLf & _
               "- at least 1 digit" & vbCrLf & _
               "- at least 1 special character", _
               vbExclamation, "Warning"  ' Show warning
        PasswordText.SetFocus  ' Focus password field for correction
        Exit Sub               ' Abort save operation
    End If
    
   ' Verify 10-digit mobile number format
    If Len(MobileNoText.Text) <> 10 Then
        MsgBox "Please enter a valid 10-digit mobile number.", vbExclamation, "Warning"  ' Show warning
        MobileNoText.SetFocus  ' Focus mobile field for correction
        Exit Sub               ' Abort save operation
    End If
    
    Dim sqlQuery As String  ' SQL command builder

    ' Build admin insertion query
    sqlQuery = "INSERT INTO admin_table(Admin_Id, Name, Password, Mobile_No, Address, Photo) values (" & _
               "'" & AdminIdText.Text & "', '" & Trim(NameText.Text) & "', '" & PasswordText.Text & "', " & _
               "'" & MobileNoText.Text & "', '" & Trim(AddressText.Text) & "', '" & imagePath & "')"
    
    On Error GoTo SaveError  ' Redirect runtime errors to SaveError handler
    
        ConnDB.Execute sqlQuery  ' Execute database command
        MsgBox "New Admin Added", vbInformation, "Message"  ' Show success notification
        disableControl           ' Deactivate input controls
        AddAdminButton.Enabled = True    ' Reactivate add button
        SaveAdminButton.Enabled = False  ' Deactivate save button
        Exit Sub                 ' Exit before error handler
    
' Save Error Handler
SaveError:
        disableControl          ' Deactivate admin data input controls
        AddAdminButton.Enabled = True    ' Reactivate add button
        SaveAdminButton.Enabled = False  ' Deactivate save button
        MsgBox "Failed to add admin: " & Err.Description, vbCritical, "Error"  ' Show error details
    
End Sub

' --- CONTROL ENABLE ROUTINE ---
' Activates form input controls

Private Sub enableControl()

    NameText.Enabled = True           ' Activate name field
    AdminIdText.Enabled = True        ' Activate admin id field
    PasswordText.Enabled = True       ' Activate password field
    MobileNoText.Enabled = True       ' Activate mobile number field
    AddressText.Enabled = True        ' Activate address field
    AdminImage.Enabled = True         ' Activate image display
    UploadPhotoButton.Enabled = True  ' Activate photo upload button

End Sub

' --- CONTROL DISABLE ROUTINE ---
' Deactivates form input controls

Private Sub disableControl()

    NameText.Enabled = False           ' Deactivate name field
    AdminIdText.Enabled = False        ' Deactivate admin id field
    PasswordText.Enabled = False       ' Deactivate password field
    MobileNoText.Enabled = False       ' Deactivate mobile number field
    AddressText.Enabled = False        ' Deactivate address field
    AdminImage.Enabled = False         ' Deactivate image display
    UploadPhotoButton.Enabled = False  ' Deactivate photo upload button

End Sub

' --- NAME TEXT KEYPRESS VALIDATION ---
' Restricts input to alphabetic characters and spaces

Private Sub NameText_KeyPress(KeyAscii As Integer)
    
    ' Allow: A-Z, a-z, space, backspace
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
        ' Valid input - allow processing
    Else
        KeyAscii = 0    ' Block invalid characters
    End If

End Sub

' --- ADMIN ID TEXT KEYPRESS VALIDATION ---
' Restricts input to alphanumeric characters

Private Sub AdminIdText_KeyPress(KeyAscii As Integer)

    ' Allow: A-Z, a-z, 0-9, backspace
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        ' Valid input - allow processing
    Else
        KeyAscii = 0  ' Block invalid characters
    End If

End Sub

' --- PASSWORD TEXT KEYPRESS VALIDATION ---
' Allows standard password characters

Private Sub PasswordText_KeyPress(KeyAscii As Integer)

    ' Allow: ASCII 33-126 (printable), backspace
    If (KeyAscii >= 33 And KeyAscii <= 126) Or KeyAscii = 8 Then
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

    isOpenAddAdmin = False  ' Reset add admin form open state flag
    
    ' Close database connection if open
    If ConnDB.State = adStateOpen Then
        ConnDB.Close    ' Close active connection
        Set ConnDB = Nothing    ' Release object memory
    End If

End Sub
