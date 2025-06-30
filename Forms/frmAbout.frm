VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10890
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10890
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      TabIndex        =   2
      Top             =   -120
      Width           =   10935
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Student  Attendance Management System"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   -1080
         TabIndex        =   3
         Top             =   600
         Width           =   12975
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   3960
      TabIndex        =   1
      Top             =   1680
      Width           =   6975
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "prashantkumarrrg777@gmail.com"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   720
         TabIndex        =   8
         Top             =   4440
         Width           =   3615
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Prashanth Kumar G"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   4200
         Width           =   2775
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   3960
         Width           =   2775
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   $"frmAbout.frx":5A5A2
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   720
         TabIndex        =   5
         Top             =   1920
         Width           =   5535
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   $"frmAbout.frx":5A647
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   5535
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   4095
      Begin VB.Image AboutImage 
         Height          =   3015
         Left            =   360
         Picture         =   "frmAbout.frx":5A6DF
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================== ABOUT SCREEN =====================
' Displays application information and credits
' Key Features:
'   - Shows project details and developer information
'   - Tracks open/closed state via flag
' ========================================================

Option Explicit

' --- GLOBAL VARIABLES ---

Public isOpenAbout As Boolean   ' About form open state flag

' --- FORM UNLOAD EVENT ---
' Handles cleanup when about screen closes

Private Sub Form_Unload(Cancel As Integer)

   isOpenAbout = False  ' Reset about form open state flag

End Sub
