VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditAdmin 
   BackColor       =   &H00400040&
   Caption         =   "Edit Admin"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16305
   Icon            =   "frmEditAdmin.frx":0000
   LinkTopic       =   "Form10"
   ScaleHeight     =   9555
   ScaleWidth      =   16305
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13080
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton DeleteAdminButton 
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9120
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox SearchAdminText 
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
      Left            =   4680
      TabIndex        =   14
      Top             =   2640
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
      Left            =   6000
      TabIndex        =   7
      Top             =   6360
      Visible         =   0   'False
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
      Left            =   6000
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
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
      Left            =   6000
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
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
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   4
      Top             =   7560
      Visible         =   0   'False
      Width           =   3495
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
      Left            =   13800
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton UpdateAdminButton 
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Visible         =   0   'False
      Width           =   2535
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Visible         =   0   'False
      Width           =   2175
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
      Left            =   9720
      TabIndex        =   22
      Top             =   4200
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
      Index           =   2
      Left            =   9720
      TabIndex        =   21
      Top             =   5400
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
      Left            =   9720
      TabIndex        =   20
      Top             =   6600
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
      Left            =   9720
      TabIndex        =   19
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
      Index           =   6
      Left            =   16200
      TabIndex        =   18
      Top             =   7080
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
      Left            =   17520
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the Admin ID To Edit"
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
      Left            =   3600
      TabIndex        =   16
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label AttributeLabel 
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
      Index           =   3
      Left            =   2400
      TabIndex        =   13
      Top             =   6480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label AttributeLabel 
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
      Index           =   2
      Left            =   2400
      TabIndex        =   12
      Top             =   5280
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
      Index           =   1
      Left            =   2400
      TabIndex        =   11
      Top             =   4080
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
      Index           =   4
      Left            =   2400
      TabIndex        =   10
      Top             =   7680
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
      Index           =   5
      Left            =   10200
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label AttributeLabel 
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
      Index           =   6
      Left            =   10200
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image AdminImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   13800
      Stretch         =   -1  'True
      Top             =   5160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Admin"
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
      Left            =   9720
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "frmEditAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================== EDIT ADMIN FORM =====================
' Interface for modifying and deleting administrator records
' Key Features:
'   - Admin record search and retrieval
'   - Selective field editing with visual indicators
'   - Secure update/delete operations
'   - Comprehensive input validation
' ==========================================================

Option Explicit

' --- GLOBAL VARIABLES ---

Dim ConnDB As New ADODB.Connection          ' Database connection object
Dim RsAdminDetails As New ADODB.Recordset   ' Admin details recordset object
Dim imagePath As String                     ' Administrator photo storage path
Dim msg As VbMsgBoxResult                   ' Message box response storage
Dim sqlQuery As String                      ' SQL command builder
Dim i As Integer                            ' Loop control variable
Public isOpenEditAdmin As Boolean           ' Edit Admin form open state flag

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
        hideControl   ' Initialize UI with controls hidden
        Exit Sub      ' Exit before error handler
    
' Database Connection Error Handler
ConnectionError:
        MsgBox "Database connection failed: " & Err.Description, vbCritical, "Error"    ' Show error details
        Unload Me   ' Terminate edit admin form
    
End Sub

' --- PROCEED SEARCH BUTTON CLICK ---
' Retrieves admin details for editing

Private Sub ProceedSearchButton_Click()

    ' Validate admin id input
    If Trim(SearchAdminText.Text) = "" Then
        clearControl     ' Reset input fields
        MsgBox "Enter Admin Id", vbExclamation, "Warning"  ' Show warning
        hideControl      ' Hide admin details section
        Exit Sub         ' Abort search operation
    End If
    
    ' Query database for admin record
    RsAdminDetails.Open "SELECT * FROM admin_table WHERE Admin_Id='" & SearchAdminText.Text & "'", ConnDB, adOpenStatic, adLockReadOnly
    
    ' Process query results
    If Not RsAdminDetails.EOF = True Then
        ' Populate form fields with admin data
        NameText.Text = RsAdminDetails.Fields("Name").Value          ' Set name field
        AdminIdText.Text = RsAdminDetails.Fields("Admin_Id").Value   ' Set admin id field
        PasswordText.Text = RsAdminDetails.Fields("Password").Value  ' Set password field
        MobileNoText.Text = RsAdminDetails.Fields("Mobile_No").Value ' Set mobile number field
        AddressText.Text = RsAdminDetails.Fields("Address").Value    ' Set address field
        imagePath = RsAdminDetails.Fields("Photo").Value             ' Store photo path
        AdminImage.Picture = LoadPicture(App.Path & "\" & RsAdminDetails.Fields("Photo").Value)  ' Set photo field
        
        showControl      ' Display admin details section
        disableControl   ' Deactivate input fields initially
        UpdateAdminButton.Enabled = False  ' Disable update functionality
        DeleteAdminButton.Enabled = True   ' Enable delete functionality
        MsgBox "Admin Found", vbInformation, "Message"  ' Show success notification
    Else
        clearControl     ' Reset input fields
        MsgBox "Invalid Admin Id", vbCritical, "Error"  ' Show error
        hideControl      ' Hide admin details section
    End If
    
    closeRsAdminDetails  ' Cleanup recordset resources

End Sub

' --- EDIT LABEL CLICK ---
' Enables editing for specific admin attribute

Private Sub EditLabel_Click(index As Integer)

    UpdateAdminButton.Enabled = True    ' Enable update functionality
    EditLabel(index).Forecolor = &HFF&  ' Change label color to red (visual indicator)
    
    ' Activate specific field based on index
    If index = 1 Then
        NameText.Enabled = True       ' Activate name field
        NameText.SetFocus             ' Focus name field for editing
    ElseIf index = 2 Then
        AdminIdText.Enabled = True    ' Activate admin id field
        AdminIdText.SetFocus          ' Focus admin id field for editing
    ElseIf index = 3 Then
        PasswordText.Enabled = True   ' Activate password field
        PasswordText.SetFocus         ' Focus password field for editing
    ElseIf index = 4 Then
        MobileNoText.Enabled = True   ' Activate mobile number field
        MobileNoText.SetFocus         ' Focus mobile number field for editing
    ElseIf index = 5 Then
        AddressText.Enabled = True    ' Activate address field
        AddressText.SetFocus          ' Focus address field for editing
    ElseIf index = 6 Then
        UploadPhotoButton.Enabled = True  ' Activate photo upload button
        UploadPhotoButton.SetFocus        ' Focus upload button for operation
    End If

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

' --- UPDATE ADMIN BUTTON CLICK ---
' Validates and saves modified admin record

Private Sub UpdateAdminButton_Click()

    ' Check for empty required fields
    If Trim(NameText.Text) = "" Or Trim(AdminIdText.Text) = "" Or Trim(PasswordText.Text) = "" Or _
       Trim(MobileNoText.Text) = "" Or Trim(AddressText.Text) = "" Or Trim(imagePath) = "" Then
            MsgBox "All Fields are Mandatory", vbExclamation, "Warning"  ' Show warning
            Exit Sub            ' Abort update operation
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
        Exit Sub               ' Abort update operation
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
        ' Build admin update query
        sqlQuery = "UPDATE admin_table SET Admin_Id='" & AdminIdText.Text & "', Name='" & Trim(NameText.Text) & "', " & _
                   "Password='" & PasswordText.Text & "', Mobile_No='" & MobileNoText.Text & "', " & _
                   "Address='" & Trim(AddressText.Text) & "', Photo='" & imagePath & "' WHERE Admin_Id='" & SearchAdminText.Text & "' "
        
        On Error GoTo UpdateError  ' Redirect runtime errors to UpdateError handler
        
            ConnDB.Execute sqlQuery  ' Execute database command
            MsgBox "Admin Information Updated", vbInformation, "Message"  ' Show success notification
            disableControl           ' Deactivate input fields
            UpdateAdminButton.Enabled = False  ' Disable update functionality
            Exit Sub                 ' Exit before error handler
            
' Update Error Handler
UpdateError:
            MsgBox "Failed to update admin: " & Err.Description, vbCritical, "Error"  ' Show error details
            disableControl          ' Deactivate input fields
            UpdateAdminButton.Enabled = False  ' Disable update functionality
    End If

End Sub

' --- DELETE ADMIN BUTTON CLICK ---
' Removes admin record from system

Private Sub DeleteAdminButton_Click()

    ' Validate admin id presence
    If Trim(SearchAdminText.Text) = "" Or Trim(AdminIdText.Text) = "" Then
        MsgBox "Admin Id is Necessary for deletion", vbExclamation, "Warning"  ' Show warning
        hideControl   ' Hide admin details section
        Exit Sub      ' Abort delete operation
    End If
    
    ' Confirm delete operation
    msg = MsgBox("Are you sure want to Delete", vbExclamation + vbYesNo, "Warning")  ' Confirmation dialog (yes/no)
    
    ' If selected yes proceed to delete
    If msg = vbYes Then
        ' Build admin deletion query
        sqlQuery = "DELETE FROM admin_table WHERE Admin_Id='" & SearchAdminText.Text & "'"
        
        On Error GoTo DeleteError  ' Redirect runtime errors to DeleteError handler
        
            ConnDB.Execute sqlQuery  ' Execute database command
            MsgBox "Admin Deleted from Records", vbInformation, "Message"  ' Show success notification
            clearControl            ' Reset input fields
            DeleteAdminButton.Enabled = False  ' Disable delete functionality
            UpdateAdminButton.Enabled = False  ' Disable update functionality
            Exit Sub                ' Exit before error handler
            
' Delete Error Handler
DeleteError:
            MsgBox "Failed to delete admin: " & Err.Description, vbCritical, "Error"  ' Show error details
            clearControl            ' Reset input fields
            DeleteAdminButton.Enabled = False  ' Disable delete functionality
    End If
    
End Sub

' --- SHOW CONTROL ROUTINE ---
' Displays admin details section

Private Sub showControl()

    ' Make attribute labels and edit controls visible
    For i = 1 To 6
        AttributeLabel(i).Visible = True  ' Show field labels
        EditLabel(i).Visible = True       ' Show edit triggers
    Next
    
    ' Make input fields and buttons visible
    NameText.Visible = True            ' Show name field
    AdminIdText.Visible = True         ' Show admin id field
    PasswordText.Visible = True        ' Show password field
    MobileNoText.Visible = True        ' Show mobile field
    AddressText.Visible = True         ' Show address field
    AdminImage.Visible = True          ' Show photo display
    UploadPhotoButton.Visible = True   ' Show upload button
    UpdateAdminButton.Visible = True   ' Show update button
    DeleteAdminButton.Visible = True   ' Show delete button

End Sub

' --- HIDE CONTROL ROUTINE ---
' Conceals admin details section

Private Sub hideControl()
    
    ' Hide attribute labels and edit controls
    For i = 1 To 6
        AttributeLabel(i).Visible = False  ' Hide field labels
        EditLabel(i).Visible = False       ' Hide edit triggers
    Next
    
    ' Hide input fields and buttons
    NameText.Visible = False            ' Hide name field
    AdminIdText.Visible = False         ' Hide admin id field
    PasswordText.Visible = False        ' Hide password field
    MobileNoText.Visible = False        ' Hide mobile field
    AddressText.Visible = False         ' Hide address field
    AdminImage.Visible = False          ' Hide photo display
    UploadPhotoButton.Visible = False   ' Hide upload button
    UpdateAdminButton.Visible = False   ' Hide update button
    DeleteAdminButton.Visible = False   ' Hide delete button

End Sub

' --- DISABLE CONTROL ROUTINE ---
' Deactivates input fields and resets edit indicators

Private Sub disableControl()

    ' Reset edit label colors to white
    For i = 1 To 6
        EditLabel(i).Forecolor = &HFFFFFF  ' Reset visual indicator
    Next
    
    ' Deactivate input fields
    NameText.Enabled = False            ' Lock name field
    AdminIdText.Enabled = False         ' Lock admin id field
    PasswordText.Enabled = False        ' Lock password field
    MobileNoText.Enabled = False        ' Lock mobile field
    AddressText.Enabled = False         ' Lock address field
    UploadPhotoButton.Enabled = False   ' Lock upload button

End Sub

' --- CLEAR CONTROL ROUTINE ---
' Resets input fields to empty state

Private Sub clearControl()

    ' Clear all input fields
    NameText.Text = ""          ' Reset name field
    AdminIdText.Text = ""       ' Reset admin id field
    PasswordText.Text = ""      ' Reset password field
    MobileNoText.Text = ""      ' Reset mobile field
    AddressText.Text = ""       ' Reset address field
    AdminImage.Picture = LoadPicture("")  ' Reset photo field

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

' --- RECORDSET CLEANUP ---
' Safely closes admin details recordset

Private Sub closeRsAdminDetails()
    
    ' Close database recordset if open
    If RsAdminDetails.State = adStateOpen Then
        RsAdminDetails.Close     ' Close active recordset
        Set RsAdminDetails = Nothing  ' Release object memory
    End If
    
End Sub

' --- FORM UNLOAD EVENT ---
' Handles cleanup when form is closed

Private Sub Form_Unload(Cancel As Integer)

    isOpenEditAdmin = False  ' Reset edit admin form open state flag
    
    ' Close database connection if open
    If ConnDB.State = adStateOpen Then
        ConnDB.Close    ' Close active connection
        Set ConnDB = Nothing    ' Release object memory
    End If

End Sub
