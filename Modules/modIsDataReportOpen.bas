Attribute VB_Name = "modIsDataReportOpen"
' ===================== REPORT STATE MANAGEMENT =====================
' Tracks open/closed status of data reporting forms
' Key Features:
'   - Centralized state flags for all reports
'   - Simple boolean status tracking
'   - Prevents duplicate report instances
' ================================================================

Option Explicit

' --- REPORT STATE FLAGS ---
' Tracks open status for each report type

Public isOpenAttendanceDetails As Boolean          ' Attendance Details report open state flag
Public isOpenOverallAttendanceDetails As Boolean   ' Overall Attendance report open state flag
Public isOpenStudentDetails As Boolean             ' Student Details report open state flag
Public isOpenAdminDetails As Boolean               ' Admin Details report open state flag
