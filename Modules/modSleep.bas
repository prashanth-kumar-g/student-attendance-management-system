Attribute VB_Name = "modSleep"
' ===================== THREAD MANAGEMENT =====================
' Provides controlled execution pausing functionality
' Key Features:
'   - Precise millisecond-level delays
'   - Windows API-based implementation
'   - Safe thread suspension
' ============================================================

Option Explicit

' --- SLEEP FUNCTION DECLARATION ---
' Suspends execution for specified duration
' Parameters:
'   dwMilliseconds: Delay duration in milliseconds (Long)

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
