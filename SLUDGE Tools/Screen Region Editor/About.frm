VERSION 5.00
Begin VB.Form AboutBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   30
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CloseButton 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label AboutLabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains the about dialog box.
Option Explicit

'This procedure closes this window.
Private Sub CloseButton_Click()
On Error GoTo ErrorTrap
   Unload Me
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   Me.Caption = "About " & App.Title
   
   With AboutLabel
      .Caption = App.Title & vbCrLf
      .Caption = .Caption & "Version: " & App.Major & "." & App.Minor & App.Revision & vbCrLf
      .Caption = .Caption & "***2006***" & vbCrLf
      .Caption = .Caption & "by: " & App.CompanyName & vbCrLf
      .Caption = .Caption & vbCrLf
      .Caption = .Caption & App.LegalCopyright
   End With
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

