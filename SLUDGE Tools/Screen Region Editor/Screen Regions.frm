VERSION 5.00
Begin VB.Form ScreenRegionsBox 
   BackColor       =   &H00000000&
   Caption         =   "Screen Regions"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   ClipControls    =   0   'False
   Icon            =   "Screen Regions.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   Begin VB.PictureBox EditingBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawMode        =   6  'Mask Pen Not
      Height          =   6135
      Left            =   0
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   565
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "ScreenRegionsBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains the screen region editing window.
Option Explicit
Private AddNewRegion As Boolean           'Indicates whether a new screen region being is added.
Private DragEditingBox As Boolean         'Indicates whether the editing box is being dragged.
Private DragX As Long                     'Contains the editing box's horizontal position while being dragged.
Private DragY As Long                     'Contains the editing box's vertical position while being dragged.
Private MoveRegion As Boolean             'Indicates whether a screen region is being moved.
Private PreviousX As Long                 'Contains the editor box's previous horizontal position.
Private PreviousY As Long                 'Contains the editor box's previous vertical position.
Private ResizeRegion As Boolean           'Indicates whether a screen region is being resized.
Private SetCharacterLocation As Boolean   'Indicates whether a character's location within a screen region is being set.
Private ShiftKeyPressed As Boolean        'Indicates whether the shift key is being pressed.

'This procedure places the editing box at the position indicated by the user by dragging.
Private Sub DropEditingBox(x As Long, y As Long)
On Error GoTo ErrorTrap
   CheckForManualCodeEditing
   
   If DragEditingBox Then
      EndDrag
      EditingBox.Move x, y
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure ends the dragging of the editing box.
Private Sub EndDrag()
On Error Resume Next
   DragEditingBox = False
   EditingBox.Drag vbEndDrag
End Sub

'This procedure indicates that new character x/y properties have been selected by the user.
Private Sub EditingBox_DblClick()
On Error GoTo ErrorTrap
   CheckForManualCodeEditing

   If Not BlockEditing Then
      If Not Selection.Region = NO_REGION Then SetCharacterLocation = True
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to place the editing box at the position indicated by the user by dragging.
Private Sub EditingBox_DragDrop(Source As Control, x As Single, y As Single)
On Error Resume Next
   If Source Is EditingBox Then DropEditingBox CLng((x - DragX) + PreviousX), CLng((y - DragY) + PreviousY)
End Sub


'This procedure gives the command to check whether the script code has been manually edited.
Private Sub EditingBox_GotFocus()
On Error GoTo ErrorTrap
   CheckForManualCodeEditing
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure processes the user's key stroke while the key is being pressed.
Private Sub EditingBox_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap
   If KeyCode = vbKeyShift Then ShiftKeyPressed = True
   
   CheckForManualCodeEditing
   If Not BlockEditing Then
      With Selection
         If Not .Region = NO_REGION Then
            GetRegionProperties .Region, , .x1, .y1, .x2, .y2
            If KeyCode = vbKeyLeft Then
               If .Corner = NoCorner Then .x1 = .x1 - Settings.GridCellWidth: .x2 = .x2 - Settings.GridCellWidth
               If .Corner = UpperLeftCorner Or .Corner = LowerLeftCorner Then .x1 = .x1 - Settings.GridCellWidth
               If .Corner = UpperRightCorner Or .Corner = LowerRightCorner Then .x2 = .x2 - Settings.GridCellWidth
            ElseIf KeyCode = vbKeyRight Then
               If .Corner = NoCorner Then .x1 = .x1 + Settings.GridCellWidth: .x2 = .x2 + Settings.GridCellWidth
               If .Corner = UpperLeftCorner Or .Corner = LowerLeftCorner Then .x1 = .x1 + Settings.GridCellWidth
               If .Corner = UpperRightCorner Or .Corner = LowerRightCorner Then .x2 = .x2 + Settings.GridCellWidth
            ElseIf KeyCode = vbKeyUp Then
               If .Corner = NoCorner Then .y1 = .y1 - Settings.GridCellHeight: .y2 = .y2 - Settings.GridCellHeight
               If .Corner = UpperLeftCorner Or .Corner = UpperRightCorner Then .y1 = .y1 - Settings.GridCellHeight
               If .Corner = LowerLeftCorner Or .Corner = LowerRightCorner Then .y2 = .y2 - Settings.GridCellHeight
            ElseIf KeyCode = vbKeyDown Then
               If .Corner = NoCorner Then .y1 = .y1 + Settings.GridCellHeight: .y2 = .y2 + Settings.GridCellHeight
               If .Corner = UpperLeftCorner Or .Corner = UpperRightCorner Then .y1 = .y1 + Settings.GridCellHeight
               If .Corner = LowerLeftCorner Or .Corner = LowerRightCorner Then .y2 = .y2 + Settings.GridCellHeight
            End If
            ChangeRegion .Region, , .x1, .y1, .x2, .y2
            DrawRegions
            DisplayProperties
         End If
      End With
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure processes the user's key strokes.
Private Sub EditingBox_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap
   CheckForManualCodeEditing

   If Not BlockEditing Then
      If KeyAscii = vbKeyTab Then
         With Selection
            If ShiftKeyPressed Then
               If .Region > 0 Then .Region = .Region - 1
            ElseIf Not ShiftKeyPressed Then
               If .Region < UBound(ScreenRegions.Properties()) - 1 Then .Region = .Region + 1
            End If
   
            DrawRegions
   
            If Not .Region = NO_REGION Then
               GetRegionProperties .Region, .ObjectType, .x1, .y1, .x2, .y2, .CharacterX, .CharacterY, .Direction
               DisplayProperties
            End If
         End With
      End If
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure processes the user's key stroke when the key is released.
Private Sub EditingBox_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap
   If KeyCode = vbKeyShift Then ShiftKeyPressed = False
   
   CheckForManualCodeEditing
      If BlockEditing Then
         BlockEditing = False
      Else
         With Selection
           If KeyCode = vbKeyReturn Then
              AddRegion LastObjectTypeUsed, EditorWidth / 4, EditorHeight / 4, EditorWidth / 1.3, EditorHeight / 1.3, , , LastDirectionUsed
              Selection.Region = UBound(ScreenRegions.Properties()) - 1
              DrawRegions
           End If
         
           If Not .Region = NO_REGION Then
              If KeyCode = vbKeyDelete Then
                 RemoveRegion .Region
                 If UBound(ScreenRegions.Properties()) = 0 Then
                    .Region = NO_REGION
                 ElseIf Not UBound(ScreenRegions.Properties()) = 0 Then
                    If Not .Region = 0 Then .Region = .Region - 1
                 End If
                 DrawRegions
                 GetRegionProperties .Region, .ObjectType, .x1, .y1, .x2, .y2, .CharacterX, .CharacterY, .Direction
                 DisplayProperties
              ElseIf KeyCode = vbKeySpace Then
                 If .Corner = LowerLeftCorner Then
                  .Corner = NoCorner
                 ElseIf Not .Corner = LowerLeftCorner Then
                  .Corner = .Corner + 1
                 End If
                 DrawRegions
              End If
           End If
         End With
      End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure checks which selected region has been selected when the user left-clicks the mouse.
Private Sub EditingBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrorTrap
   CheckForManualCodeEditing
   If Not BlockEditing Then
      If PropertiesChanged Then
         UpdateRegion
         PropertiesChanged = False
      End If
       
      If (ShiftKeyPressed And Settings.UseShiftToSelect) Or (Not ShiftKeyPressed And Not Settings.UseShiftToSelect) Then
         With Selection
            DetermineSelectedCorner CLng(x), CLng(y)
            If .Region = NO_REGION Then DetermineSelectedRegion CLng(x), CLng(y)
            If Not .Region = NO_REGION Then
               GetRegionProperties .Region, .ObjectType, .x1, .y1, .x2, .y2, .CharacterX, .CharacterY, .Direction
               DisplayProperties
            End If
         End With
      End If
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure adds/moves/resizes objects when the user moves the mouse.
Private Sub EditingBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrorTrap
Static MoveX1 As Long
Static MoveX2 As Long
Static MoveY1 As Long
Static MoveY2 As Long
Static PreviousX1 As Long
Static PreviousX2 As Long
Static PreviousY1 As Long
Static PreviousY2 As Long

   CheckForManualCodeEditing

   If Not BlockEditing Then
      With Selection
         If Button = vbLeftButton Then
            If .Region = NO_REGION And .Corner = NoCorner Then
               If AddNewRegion Then
                  EditingBox.Line (.x1, .y1)-(PreviousX2, PreviousY2), , B
                  PreviousX2 = x
                  PreviousY2 = y
                  .x2 = x
                  .y2 = y
                  EditingBox.Line (.x1, .y1)-(x, y), , B
               ElseIf Not AddNewRegion Then
                  If (ShiftKeyPressed And Settings.UseShiftToAdd) Or (Not ShiftKeyPressed And Not Settings.UseShiftToAdd) Then
                     AddNewRegion = True
                     PreviousX2 = x
                     PreviousY2 = y
                     .x1 = x
                     .y1 = y
                     EditingBox.PSet (.x1, .y1)
                  End If
               End If
            ElseIf .Corner = NoCorner Then
               Selection.Corner = NO_REGION
               Selection.Region = NO_REGION
            End If
            If Not .Region = NO_REGION Then
               If Not .Corner = NoCorner Then
                  If ResizeRegion Then
                     If .Corner = UpperLeftCorner Then
                        EditingBox.Line (PreviousX1, PreviousY1)-(.x2, .y2), , B
                        PreviousX1 = x
                        PreviousY1 = y
                        .x1 = x
                        .y1 = y
                        EditingBox.Line (.x1, .y1)-(.x2, .y2), , B
                     ElseIf .Corner = UpperRightCorner Then
                        EditingBox.Line (.x1, PreviousY1)-(PreviousX2, .y2), , B
                        PreviousX2 = x
                        PreviousY1 = y
                        .x2 = x
                        .y1 = y
                        EditingBox.Line (.x1, .y1)-(.x2, .y2), , B
                     ElseIf .Corner = LowerRightCorner Then
                        EditingBox.Line (.x1, .y1)-(PreviousX2, PreviousY2), , B
                        PreviousX2 = x
                        PreviousY2 = y
                        .x2 = x
                        .y2 = y
                        EditingBox.Line (.x1, .y1)-(.x2, .y2), , B
                     ElseIf .Corner = LowerLeftCorner Then
                        EditingBox.Line (PreviousX1, .y1)-(.x2, PreviousY2), , B
                        PreviousX1 = x
                        PreviousY2 = y
                        .x1 = x
                        .y2 = y
                        EditingBox.Line (.x1, .y1)-(.x2, .y2), , B
                     End If
                  ElseIf Not ResizeRegion Then
                     ResizeRegion = True
                     If .Corner = UpperLeftCorner Then
                        PreviousX1 = .x1
                        PreviousY1 = .y1
                     ElseIf .Corner = UpperRightCorner Then
                        PreviousX2 = .x2
                        PreviousY1 = .y1
                     ElseIf .Corner = LowerRightCorner Then
                        PreviousX2 = .x2
                        PreviousY2 = .y2
                     ElseIf .Corner = LowerLeftCorner Then
                        PreviousX1 = .x1
                        PreviousY2 = .y2
                     End If
                  End If
               End If
            End If
         ElseIf Button = vbMiddleButton Then
            If .Corner = NoCorner Then
               If MoveRegion Then
                  EditingBox.Line (PreviousX1, PreviousY1)-(PreviousX2, PreviousY2), , B
                  PreviousX1 = .x1
                  PreviousY1 = .y1
                  PreviousX2 = .x2
                  PreviousY2 = .y2
                  .x1 = x - MoveX1
                  .y1 = y - MoveY1
                  .x2 = x + MoveX2
                  .y2 = y + MoveY2
                  EditingBox.Line (.x1, .y1)-(.x2, .y2), , B
               ElseIf Not MoveRegion Then
                  MoveRegion = True
                  MoveX1 = x - .x1
                  MoveY1 = y - .y1
                  MoveX2 = .x2 - x
                  MoveY2 = .y2 - y
                  PreviousX1 = .x1
                  PreviousY1 = .y1
                  PreviousX2 = .x2
                  PreviousY2 = .y2
               End If
            End If
         ElseIf Button = vbRightButton Then
            If Not DragEditingBox Then
               DragEditingBox = True
               DragX = x
               DragY = y
               PreviousX = EditingBox.Left
               PreviousY = EditingBox.Top
               EditingBox.Drag vbBeginDrag
            End If
         End If
      
         If AddNewRegion Or MoveRegion Or ResizeRegion Then
            Me.Caption = "Screen Regions - " & .x1 & ", " & .y1 & " - " & x & ", " & y & " = " & Abs(.x1 - .x2) & " x " & Abs(.y1 - .y2)
         ElseIf Not (AddNewRegion Or ResizeRegion) Then
            Me.Caption = "Screen Regions - " & x & ", " & y
         End If
      End With
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure finishes adding/moving/resizing a screen region.
Private Sub EditingBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrorTrap
   CheckForManualCodeEditing

   If BlockEditing Then
      BlockEditing = False
   Else
      With Selection
         If Button = vbLeftButton Then
            If AddNewRegion Then
               AddNewRegion = False
               AddRegion LastObjectTypeUsed, .x1, .y1, .x2, .y2, .CharacterX, .CharacterY, LastDirectionUsed
               DetermineSelectedRegion CLng(x), CLng(y)
            ElseIf ResizeRegion Then
               ResizeRegion = False
               ChangeRegion .Region, , .x1, .y1, .x2, .y2, .CharacterX, .CharacterY
            End If
            If SetCharacterLocation Then
               SetCharacterLocation = False
               .CharacterX = x
               .CharacterY = y
               ChangeRegion .Region, , , , , , .CharacterX, .CharacterY
            End If
         ElseIf Button = vbMiddleButton Then
            If MoveRegion Then
                MoveRegion = False
                ChangeRegion .Region, , .x1, .y1, .x2, .y2, .CharacterX, .CharacterY
            End If
         End If
         DrawRegions
         GetRegionProperties .Region, .ObjectType, .x1, .y1, .x2, .y2, .CharacterX, .CharacterY, .Direction
         DisplayProperties
         EndDrag
      End With
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to check whether the script code has been manually edited.
Private Sub Form_Activate()
On Error GoTo ErrorTrap
   CheckForManualCodeEditing
   DrawRegions
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to place the editing box at the position indicated by the user by dragging.
Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
On Error Resume Next
   If Source Is EditingBox Then DropEditingBox CLng(x - DragX), CLng(y - DragY)
End Sub

'This procedure gives the command to check whether the script code has been manually edited.
Private Sub Form_GotFocus()
On Error GoTo ErrorTrap
   CheckForManualCodeEditing
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error Resume Next
   Me.Left = 0
   Me.Top = 0
   Me.Width = (EditorWidth + 20) * Screen.TwipsPerPixelX
   Me.Height = (EditorHeight + 36) * Screen.TwipsPerPixelY
   
   AddNewRegion = False
   DragEditingBox = False
   MoveRegion = False
   ResizeRegion = False
   Selection.Corner = NoCorner
   Selection.Region = NO_REGION
   ShiftKeyPressed = False
   
   If EditorWidth = 0 Then
      EditingBox.Width = Me.ScaleWidth
   ElseIf Not EditorWidth = 0 Then
      EditingBox.Width = EditorWidth + 4
   End If
   
   If EditorHeight = 0 Then
      EditingBox.Height = Me.ScaleHeight
   ElseIf Not EditorHeight = 0 Then
      EditingBox.Height = EditorHeight + 4
   End If
End Sub

