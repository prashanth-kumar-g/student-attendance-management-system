VERSION 5.00
Begin VB.Form frmAdminLogin 
   BackColor       =   &H00004000&
   Caption         =   "Login Details"
   ClientHeight    =   5715
   ClientLeft      =   5670
   ClientTop       =   3435
   ClientWidth     =   10905
   Icon            =   "frmAdminLogin.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10905
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000000FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton LoginButton 
      BackColor       =   &H0000FF00&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox PasswordText 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5760
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox AdminText 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Image AdminLoginImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   1545
      Left            =   1680
      Picture         =   "frmAdminLogin.frx":5A5A2
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1365
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Id"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Login"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   3720
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "frmAdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================== ADMIN LOGIN FORM =====================
' Handles administrator authentication
' Key Features:
'   - Database connection management
'   - Credential validation
'   - Secure session initialization
'   - Error handling for DB operations
' ===========================================================

Option Explicit

' --- GLOBAL VARIABLES ---

Dim ConnDB As New ADODB.Connection  ' Database connection object
Dim RsAdminLogin As New ADODB.Recordset     ' Database recordset object

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
        Unload Me   ' Terminate admin login form
    
End Sub

' --- LOGIN BUTTON CLICK ---
' Validates and authenticates administrator credentials

Private Sub LoginButton_Click()
    
    ' Check for empty credentials and provide specific feedback
    If Trim(AdminText.Text) = "" And Trim(PasswordText.Text) = "" Then
        MsgBox "Enter Admin Id and Password", vbExclamation, "Warning"  ' Both fields empty
        AdminText.SetFocus  ' Focus admin id field for input
        Exit Sub            ' Abort login process
    ElseIf Trim(AdminText.Text) <> "" And Trim(PasswordText.Text) = "" Then
        MsgBox "Enter Password", vbExclamation, "Warning"  ' Missing password
        PasswordText.SetFocus  ' Focus password field for input
        Exit Sub               ' Abort login process
    ElseIf Trim(AdminText.Text) = "" And Trim(PasswordText.Text) <> "" Then
        MsgBox "Enter Admin Id", vbExclamation, "Warning"  ' Missing admin id
        AdminText.SetFocus  ' Focus admin id field for input
        Exit Sub            ' Abort login process
    End If
    
    On Error GoTo LoginError    ' Redirect runtime errors to LoginError handler
    
        ' Query database for admin credentials
        RsAdminLogin.Open "SELECT Admin_Id, Password FROM admin_table WHERE Admin_Id = '" & Trim(AdminText.Text) & "' ", ConnDB, adOpenStatic, adLockReadOnly
        
        ' Check if admin id exists
        If Not RsAdminLogin.EOF = True Then
            ' Verify password match
            If Trim(AdminText.Text) = RsAdminLogin.Fields("Admin_Id").Value And Trim(PasswordText.Text) = RsAdminLogin.Fields("Password").Value Then
                ' Successful login sequence
                mdiMainHome.BackgroundCoverPicture.Visible = False  ' Hide default background of main MDI form
                mdiMainHome.Enabled = True                          ' Enable main MDI interface
                frmAdminPanel.Show                                  ' Launch admin panel interface
                frmAdminPanel.Enabled = True                        ' Enable admin panel controls
                frmAdminPanel.MarkAttendancePicture.SetFocus        ' Set focus to picture for default view
                MsgBox "Login Successful", vbInformation, "Message" ' Notify user
                closeRsAdminLogin    ' Cleanup recordset
                Unload Me            ' Terminate admin login form
            Else
                ' Password mismatch
                MsgBox "Invalid Password", vbCritical, "Error"  ' Show error details
                PasswordText.Text = ""     ' Clear password field
                PasswordText.SetFocus      ' Refocus for correction
            End If
        Else
            ' Admin Id not found
            MsgBox "Invalid Admin Id", vbCritical, "Error"  ' Show error details
            AdminText.Text = ""      ' Clear admin id field
            AdminText.SetFocus       ' Refocus for correction
        End If
        
        closeRsAdminLogin   ' Cleanup recordset
        Exit Sub            ' Exit before error handler
    
' Login Error Handler
LoginError:
        MsgBox "Failed to login: " & Err.Description, vbCritical, "Error"   ' Show error details
        closeRsAdminLogin   ' Ensure recordset cleanup

End Sub

' --- CANCEL BUTTON CLICK ---
' Aborts login process and returns to main interface

Private Sub CancelButton_Click()

    mdiMainHome.Enabled = True  ' Re-enable main MDI interface
    mdiMainHome.Show            ' Show main MDI interface
    Unload Me                   ' Terminate admin login form
    
End Sub

' --- RECORDSET CLEANUP ---
' Safely closes admin login recordset

Private Sub closeRsAdminLogin()

    ' Close database recordset if open
    If RsAdminLogin.State = adStateOpen Then
        RsAdminLogin.Close     ' Close active recordset
        Set RsAdminLogin = Nothing  ' Release object memory
    End If
    
End Sub

' --- FORM UNLOAD EVENT ---
' Cleans up resources when form closes

Private Sub Form_Unload(Cancel As Integer)

    ' Close database connection if open
    If ConnDB.State = adStateOpen Then
        ConnDB.Close    ' Close active connection
        Set ConnDB = Nothing    ' Release object memory
    End If
    
    ' Reset main interface state
    mdiMainHome.Enabled = True  ' Re-enable main MDI interface
    mdiMainHome.Show    ' Show main MDI interface

End Sub
