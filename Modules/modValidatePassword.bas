Attribute VB_Name = "modValidatePassword"
' ===================== PASSWORD VALIDATION =====================
' Enforces strong password security policies
' Key Features:
'   - Minimum 8-character length requirement
'   - Mandatory character diversity checks
'   - Regular expression-based validation
' ===============================================================

Option Explicit

' --- PASSWORD VALIDATION ROUTINE ---
' Verifies password meets complexity requirements
' Parameters:
'   password: String to validate
' Returns: Boolean (True = valid, False = invalid)
' Complexity Rules:
'   - At least 8 characters
'   - At least 1 lowercase letter
'   - At least 1 uppercase letter
'   - At least 1 digit
'   - At least 1 special characte

Public Function validatePassword(password As String) As Boolean

    Dim regEx As Object  ' Regular expression object
    
    ' Minimum length check
    If Len(password) < 8 Then
        validatePassword = False  ' Fail: insufficient length
        Exit Function
    End If
    
    ' Initialize regular expression engine
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Configure regex pattern for complexity:
    regEx.Pattern = "^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*\W).+$"  ' Requires lowercase, uppercase, digit, special char
    regEx.IgnoreCase = False  ' Case-sensitive matching
    regEx.Global = False      ' Single match optimization
    
    ' Execute pattern matching
    validatePassword = regEx.Test(password)  ' True if all requirements met
    
    ' Cleanup regex resources
    Set regEx = Nothing  ' Release COM object

End Function
