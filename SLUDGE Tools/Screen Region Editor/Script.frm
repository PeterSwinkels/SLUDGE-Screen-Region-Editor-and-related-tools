VERSION 5.00
Begin VB.Form ScriptBox 
   Caption         =   "Script"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "Script.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   13.125
   ScaleMode       =   4  'Character
   ScaleWidth      =   39
   Begin VB.TextBox CodeBox 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      MaxLength       =   65535
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "ScriptBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains the script code editing window.
Option Explicit

'This procedure updates the script code when the code box's text is changed.
Private Sub CodeBox_Change()
On Error GoTo ErrorTrap
   With ScriptCode
      .ManuallyEdited = (GetFocus() = CodeBox.hWnd)
      .NotSaved = True
      .Code = CodeBox.Text
   End With

   If Len(CodeBox.Text) >= CodeBox.MaxLength Then
      MsgBox "The script code has reached it maximum allowed size of " & CStr(CodeBox.MaxLength) & " bytes.", vbExclamation
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to load the first file of a set dragged and dropped into the code box.
Private Sub CodeBox_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrorTrap
   If Data.Files.Count > 0 Then
      If ScriptCode.NotSaved Then
         If MsgBox("Save the current script first?", vbYesNo Or vbQuestion Or vbDefaultButton1) = vbYes Then SaveCurrentScript
      End If
      StartNewScript
      LoadScript Data.Files(1)
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure hides this window if it should not be visible.
Private Sub Form_Activate()
On Error GoTo ErrorTrap
   If ScriptBoxVisible Then
      SSREBox.EditMenu.Visible = True
   ElseIf Not ScriptBoxVisible Then
      ScriptCode.ManuallyEdited = False
      Me.Hide
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure hides the Edit menu when this window looses the focus.
Private Sub Form_Deactivate()
On Error Resume Next
   SSREBox.EditMenu.Visible = False
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error Resume Next
Dim OldCodeManuallyEdited As Boolean
Dim OldCodeNotSaved As Boolean

   Me.Width = (SSREBox.ScaleWidth / 1.1)
   Me.Height = (SSREBox.ScaleHeight / 1.1)
   
   Me.Left = 0
   Me.Top = 0

   With ScriptCode
      OldCodeManuallyEdited = .ManuallyEdited
      OldCodeNotSaved = .NotSaved
      CodeBox.Text = .Code
      .ManuallyEdited = OldCodeManuallyEdited
      .NotSaved = OldCodeNotSaved
   End With
End Sub

'This procedure adjusts the size of the interface elements to the new window size.
Private Sub Form_Resize()
On Error Resume Next
   CodeBox.Width = Me.ScaleWidth
   CodeBox.Height = Me.ScaleHeight
End Sub

'This procedure minimizes this window when it is closed and the program is not quitting.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
   ScriptBoxVisible = False
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

