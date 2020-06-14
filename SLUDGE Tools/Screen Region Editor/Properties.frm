VERSION 5.00
Begin VB.Form PropertiesBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3345
   ClipControls    =   0   'False
   Icon            =   "Properties.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   15.563
   ScaleMode       =   4  'Character
   ScaleWidth      =   27.875
   Begin VB.ComboBox CharacterXYBox 
      Height          =   315
      ItemData        =   "Properties.frx":014A
      Left            =   1320
      List            =   "Properties.frx":016C
      TabIndex        =   7
      Text            =   "User defined."
      ToolTipText     =   "Changes the character x/y properties to a predefined value."
      Top             =   2640
      Width           =   1935
   End
   Begin VB.ComboBox DirectionBox 
      Height          =   315
      ItemData        =   "Properties.frx":021C
      Left            =   1320
      List            =   "Properties.frx":0238
      TabIndex        =   8
      ToolTipText     =   "The direction a character will face when forced."
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox RankBox 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "The rank of the screen region within the selection."
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox CharacterYBox 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   6
      Text            =   "0"
      ToolTipText     =   "The vertical position at which a character will appear when forced."
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox CharacterXBox 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   5
      Text            =   "0"
      ToolTipText     =   "The horizontal position at which a character will appear when forced."
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox y2Box 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "The vertical position of the lower right corner."
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox x2Box 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "The horizontal position of the lower right corner."
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox y1Box 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "The vertical position of the upper left corner."
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox x1Box 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "The horizontal position of the upper left corner."
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox ObjectTypeBox 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "The objecttype represented by the screen region."
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label CharacterXYLabel 
      AutoSize        =   -1  'True
      Caption         =   "Character xy:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   1140
   End
   Begin VB.Label RankLabel 
      AutoSize        =   -1  'True
      Caption         =   "Rank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   18
      Top             =   3360
      Width           =   525
   End
   Begin VB.Label DirectionLabel 
      AutoSize        =   -1  'True
      Caption         =   "Direction:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   3000
      Width           =   840
   End
   Begin VB.Label yLabel 
      AutoSize        =   -1  'True
      Caption         =   "Character y:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   1050
   End
   Begin VB.Label xLabel 
      AutoSize        =   -1  'True
      Caption         =   "Character x:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   1050
   End
   Begin VB.Label y2Label 
      AutoSize        =   -1  'True
      Caption         =   "y2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   14
      Top             =   1560
      Width           =   270
   End
   Begin VB.Label x2Label 
      AutoSize        =   -1  'True
      Caption         =   "x2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   13
      Top             =   1200
      Width           =   270
   End
   Begin VB.Label y1Label 
      AutoSize        =   -1  'True
      Caption         =   "y1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   12
      Top             =   840
      Width           =   270
   End
   Begin VB.Label x1Label 
      AutoSize        =   -1  'True
      Caption         =   "x1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   11
      Top             =   480
      Width           =   270
   End
   Begin VB.Label ObjectTypeLabel 
      AutoSize        =   -1  'True
      Caption         =   "ObjectType:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "PropertiesBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains the screen region properties window.
Option Explicit

'This procedure automatically selects the property data when the property field is selected.
Private Sub CharacterXBox_GotFocus()
On Error GoTo ErrorTrap
   CharacterXBox.SelStart = 0
   CharacterXBox.SelLength = Len(CharacterXBox.Text)
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure updates the selected screen region's Character x property.
Private Sub CharacterXBox_LostFocus()
On Error GoTo ErrorTrap
   UpdateRegion
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to change the character x/y properties to a predefined value selected from the list by the user.
Private Sub CharacterXYBox_Click()
On Error GoTo ErrorTrap
   SetCharacterXY CharacterXYBox.ListIndex
   CharacterXBox.Text = Selection.CharacterX
   CharacterYBox.Text = Selection.CharacterY
   SelectedCharacterXY = CharacterXYBox.ListIndex
   UpdateRegion
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to change the character x/y properties to a predefined value manually entered by the user.
Private Sub CharacterXYBox_LostFocus()
On Error GoTo ErrorTrap
Dim Index As Long
Dim NewCharacterXY As String
   
   NewCharacterXY = UCase$(Trim$(CharacterXYBox.Text))
   If Not Right$(NewCharacterXY, 1) = "." Then NewCharacterXY = NewCharacterXY & "."
     
   For Index = 0 To CharacterXYBox.ListCount - 1
      If UCase$(CharacterXYBox.List(Index)) = NewCharacterXY Then
         CharacterXYBox.ListIndex = Index
      End If
   Next Index
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure automatically selects the property data when the property field is selected.
Private Sub CharacterYBox_GotFocus()
On Error GoTo ErrorTrap
   CharacterYBox.SelStart = 0
   CharacterYBox.SelLength = Len(CharacterYBox.Text)
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure updates the selected screen region's Character y property.
Private Sub CharacterYBox_LostFocus()
On Error GoTo ErrorTrap
   UpdateRegion
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to update the selected screen region.
Private Sub DirectionBox_Click()
On Error GoTo ErrorTrap
   UpdateRegion
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure updates the selected screen region's Direction property.
Private Sub DirectionBox_LostFocus()
On Error GoTo ErrorTrap
   DirectionBox.Text = UCase$(Trim$(DirectionBox.Text))
   UpdateRegion
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure marks the properties as having been changed.
Private Sub Form_Activate()
On Error GoTo ErrorTrap
   PropertiesChanged = True
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error Resume Next
   Me.Left = (SSREBox.ScaleWidth / 1.01) - Me.Width
   Me.Top = Screen.TwipsPerPixelX ^ 2
   
   PropertiesBoxVisible = True
End Sub

'This procedure hides this window.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
   PropertiesBoxVisible = False
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure automatically selects the property data when the property field is selected.
Private Sub ObjectTypeBox_GotFocus()
On Error GoTo ErrorTrap
   ObjectTypeBox.SelStart = 0
   ObjectTypeBox.SelLength = Len(ObjectTypeBox.Text)
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure updates the selected screen region's ObjectType property.
Private Sub ObjectTypeBox_LostFocus()
On Error GoTo ErrorTrap
   UpdateRegion
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure automatically selects the property data when the property field is selected.
Private Sub x1Box_GotFocus()
On Error GoTo ErrorTrap
   x1Box.SelStart = 0
   x1Box.SelLength = Len(x1Box.Text)
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure updates the selected screen region's x1 property.
Private Sub x1Box_LostFocus()
On Error GoTo ErrorTrap
   UpdateRegion
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure automatically selects the property data when the property field is selected.
Private Sub x2Box_GotFocus()
On Error GoTo ErrorTrap
   x2Box.SelStart = 0
   x2Box.SelLength = Len(x2Box.Text)
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure updates the selected screen region's x2 property.
Private Sub x2Box_LostFocus()
On Error GoTo ErrorTrap
   UpdateRegion
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure automatically selects the property data when the property field is selected.
Private Sub y1Box_GotFocus()
On Error GoTo ErrorTrap
   y1Box.SelStart = 0
   y1Box.SelLength = Len(y1Box.Text)
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure updates the selected screen region's y1 property.
Private Sub y1Box_LostFocus()
On Error GoTo ErrorTrap
   UpdateRegion
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure automatically selects the property data when the property field is selected.
Private Sub y2Box_GotFocus()
On Error GoTo ErrorTrap
   y2Box.SelStart = 0
   y2Box.SelLength = Len(y2Box.Text)
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure updates the selected screen region's y2 property.
Private Sub y2Box_LostFocus()
On Error GoTo ErrorTrap
   UpdateRegion
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

