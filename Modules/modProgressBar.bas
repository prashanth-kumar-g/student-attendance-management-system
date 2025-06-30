Attribute VB_Name = "modProgressBar"
' ===================== PROGRESS BAR CUSTOMIZATION =====================
' Provides advanced styling for VB6 progress bars
' Key Features:
'   - Custom background and fill colors
'   - Win32 API-based implementation
'   - Hardware-accelerated rendering
' =====================================================================

Option Explicit

' --- WINDOWS API DECLARATIONS ---
' External functions for progress bar customization

' Purpose: Retrieves information about the specified window
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
' Purpose: Sets information for the specified window
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
' Purpose: Sends a message to a window procedure
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' --- MESSAGE CONSTANTS ---
' Windows message identifiers for progress bar control

Private Const WM_USER = &H400                          ' Base for user-defined messages
Private Const PBM_SETBARCOLOR = (WM_USER + 9)          ' Sets progress bar fill color
Private Const CCM_FIRST = &H2000                       ' Common control messages base
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)         ' Sets control background color

' --- PROGRESS BAR COLORING ROUTINE ---
' Applies custom colors to progress bar controls
' Parameters:
'   PB: ProgressBar control instance
'   Backcolor: Background color (Long)
'   Forecolor: Fill color (Long)

Public Sub PBcolor(PB As ProgressBar, Backcolor As Long, Forecolor As Long)

    SendMessage PB.hwnd, CCM_SETBKCOLOR, 0, ByVal Backcolor     ' Set background color (area behind progress)
    
    SendMessage PB.hwnd, PBM_SETBARCOLOR, 0, ByVal Forecolor    ' Set foreground color (progress fill)

End Sub
