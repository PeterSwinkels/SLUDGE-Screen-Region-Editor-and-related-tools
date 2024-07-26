VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm SSREBox 
   BackColor       =   &H8000000C&
   Caption         =   "SLUDGE Screen Region Editor"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7770
   Icon            =   "SSRE.frx":0000
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog DialogBox 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu ScriptMenu 
      Caption         =   "Sc&ript"
      Begin VB.Menu NewScriptMenu 
         Caption         =   "&New Script"
         Shortcut        =   ^N
      End
      Begin VB.Menu Separator1ScriptMenu 
         Caption         =   "-"
      End
      Begin VB.Menu LoadScriptMenu 
         Caption         =   "&Load Script"
         Shortcut        =   ^L
      End
      Begin VB.Menu Separator2ScriptMenu 
         Caption         =   "-"
      End
      Begin VB.Menu SaveScriptMenu 
         Caption         =   "&Save Script"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveScriptAsMenu 
         Caption         =   "Save Script &As"
      End
      Begin VB.Menu Separator3ScriptMenu 
         Caption         =   "-"
      End
      Begin VB.Menu CloseScriptMenu 
         Caption         =   "&Close Script"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu CutMenu 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu CopyMenu 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Separator1EditMenu 
         Caption         =   "-"
      End
      Begin VB.Menu PasteMenu 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu SelectAllMenu 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu Separator2EditMenu 
         Caption         =   "-"
      End
      Begin VB.Menu FindMenu 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu FindNextMenu 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Separator3EditMenu 
         Caption         =   "-"
      End
      Begin VB.Menu ReplaceMenu 
         Caption         =   "&Replace"
         Shortcut        =   ^H
      End
      Begin VB.Menu ReplaceNextMenu 
         Caption         =   "Replace N&ext"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu ScreenRegionsMenu 
      Caption         =   "&Screen Regions"
      Begin VB.Menu GetFromCodeSelectionMenu 
         Caption         =   "&Get From Code Selection"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu SettingsMenu 
      Caption         =   "Se&ttings"
      Begin VB.Menu EditingFieldMenuSize 
         Caption         =   "&Editing Field Size"
      End
      Begin VB.Menu GridCellSizeMenu 
         Caption         =   "&Grid Cell Size"
      End
      Begin VB.Menu Separator1SettingsMenu 
         Caption         =   "-"
      End
      Begin VB.Menu ScreenRegionDisplayColorMenu 
         Caption         =   "Screen Region Display &Color"
         Begin VB.Menu InvertedMenu 
            Caption         =   "&Inverted"
         End
         Begin VB.Menu OtherMenu 
            Caption         =   "&Other"
         End
      End
      Begin VB.Menu ScreenRegionHandleSizeMenu 
         Caption         =   "Screen Region &Handle Size"
      End
      Begin VB.Menu Separator2SettingsMenu 
         Caption         =   "-"
      End
      Begin VB.Menu UseSeparateTextLinesMenu 
         Caption         =   "Use Separate Text &Lines"
      End
      Begin VB.Menu UseShiftToAddMenu 
         Caption         =   "Use Shift To &Add"
      End
      Begin VB.Menu UseShiftToSelectMenu 
         Caption         =   "Use Shift To &Select"
      End
   End
   Begin VB.Menu ImageMenu 
      Caption         =   "&Image"
      Begin VB.Menu LoadImageMenu 
         Caption         =   "&Load Image"
         Shortcut        =   ^I
      End
      Begin VB.Menu Separator1ImageMenu 
         Caption         =   "-"
      End
      Begin VB.Menu RemoveImageMenu 
         Caption         =   "&Remove Image"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu WindowMainMenu 
      Caption         =   "&Window"
      Index           =   0
      Begin VB.Menu WindowMenu 
         Caption         =   "Screen Region &Properties"
         Index           =   0
         Shortcut        =   ^P
      End
      Begin VB.Menu WindowMenu 
         Caption         =   "Screen &Regions"
         Index           =   1
         Shortcut        =   ^R
      End
      Begin VB.Menu WindowMenu 
         Caption         =   "&Script"
         Index           =   2
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu InformationMenu 
      Caption         =   "I&nformation"
      Begin VB.Menu HelpMenu 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Separator1InformationMenu 
         Caption         =   "-"
      End
      Begin VB.Menu AboutMenu 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "SSREBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main window.
Option Explicit

'This procedure displays the about dialog box.
Private Sub AboutMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   
   AboutBox.Show vbModal
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to quit this program after user confirmation.
Private Sub CloseScriptMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   Unload Me
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure copies the selected code from the script code.
Private Sub CopyMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   
   Clipboard.SetText ScriptBox.CodeBox.SelText
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure cuts the selected code from the script code.
Private Sub CutMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   
   Clipboard.SetText ScriptBox.CodeBox.SelText
   ScriptBox.CodeBox.SelText = vbNullString
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure requests the user to specify the editing window's dimensions.
Private Sub EditingFieldMenuSize_Click()
On Error GoTo ErrorTrap
Dim NewEditorHeight As Long
Dim NewEditorWidth As Long

   With ScreenRegionsBox
      NewEditorWidth = CLng(Val(InputBox$("Editing window width (in pixels.)", , CStr(.EditingBox.Width - 4))))
      NewEditorHeight = CLng(Val(InputBox$("Editing window height (in pixels.)", , CStr(.EditingBox.Height - 4))))
      
      If Not (NewEditorWidth = 0 Or NewEditorHeight = 0) Then
         EditorWidth = NewEditorWidth
         EditorHeight = NewEditorHeight
         If Not .WindowState = vbMaximized Then
            .Width = (EditorWidth + 20) * Screen.TwipsPerPixelX
            .Height = (EditorHeight + 36) * Screen.TwipsPerPixelY
         End If
         .EditingBox.Width = EditorWidth + 4
         .EditingBox.Height = EditorHeight + 4
         If Not (.EditingBox.Width = EditorWidth + 4 Or .EditingBox.Height = EditorHeight + 4) Then
            MsgBox "The specified values are out" & vbCrLf & "of range and have been adjusted.", vbInformation
         End If
      End If
   End With
   
   DrawRegions
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to search for the specified text.
Private Sub FindMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   
   SearchForText Find
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to search for the next occurrence of the specified text.
Private Sub FindNextMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   
   SearchForText FindNext
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gets the screen regions from the selected code.
Private Sub GetFromCodeSelectionMenu_Click()
On Error Resume Next
   BlockEditing = True
   
   If ScriptBox.CodeBox.SelLength = 0 Then
      GetRegionsFromCode ScriptBox.CodeBox.Text, 0
   ElseIf Not ScriptBox.CodeBox.SelLength = 0 Then
      GetRegionsFromCode ScriptBox.CodeBox.SelText, ScriptBox.CodeBox.SelStart
   End If
   
   ScriptCode.ManuallyEdited = False
   Selection.Region = NO_REGION
   
   DrawRegions
   ScreenRegionsBox.Show
   ScreenRegionsBox.ZOrder
   ScreenRegionsBox.WindowState = vbNormal
End Sub

'This procedure requests the user to specify the screen region editing window grid cell dimensions.
Private Sub GridCellSizeMenu_Click()
On Error GoTo ErrorTrap
Dim NewGridCellHeight As Long
Dim NewGridCellWidth As Long
   
   With Settings
      BlockEditing = True
      
      NewGridCellWidth = CLng(Val(InputBox$("Grid cell width 1-100 (in pixels.)", , CStr(.GridCellWidth))))
      If NewGridCellWidth > 0 And NewGridCellWidth < 101 Then
         .GridCellWidth = NewGridCellWidth
      ElseIf Not NewGridCellWidth = 0 Then
         MsgBox "The specified grid cell width is out of range.", vbExclamation
      End If
      
      NewGridCellHeight = CLng(Val(InputBox$("Grid cell height 1-100 (in pixels.)", , CStr(.GridCellHeight))))
      If NewGridCellHeight > 0 And NewGridCellHeight < 101 Then
         .GridCellHeight = NewGridCellHeight
      ElseIf Not NewGridCellHeight = 0 Then
         MsgBox "The specified grid cell height is out of range.", vbExclamation
      End If
   End With
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure displays the help.
Private Sub HelpMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   
   If Dir$(ApplicationPath() & "Help.hta") = vbNullString Then
      MsgBox "Could not find the help.", vbExclamation
   Else
      Shell "Mshta.exe """ & ApplicationPath & "Help.hta""", vbMaximizedFocus
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure toggles the display inverted colors option on/off.
Private Sub InvertedMenu_Click()
On Error GoTo ErrorTrap
   With Settings
      BlockEditing = True
      
      .DisplayInverted = Not .DisplayInverted
      InvertedMenu.Checked = .DisplayInverted
      OtherMenu.Checked = Not .DisplayInverted
      DrawRegions
   End With
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure loads and displays the specified image.
Private Sub LoadImageMenu_Click()
On Error GoTo ErrorTrap
Dim TGALoader As New TGALoaderClass

   BlockEditing = True
   
   With DialogBox
      .DialogTitle = "Load Image"
      .FileName = vbNullString
      .Filter = "Truevision Targa (*.tga)|*.tga|"
      .Filter = .Filter & "Bitmap (*.bmp)|*.bmp|"
      .Filter = .Filter & "Icon (*.ico)|*.ico|"
      .Filter = .Filter & "Run-Length Encoded files (*.rle)|*.rle|"
      .Filter = .Filter & "Metafile (*.wmf)|*.wmf|"
      .Filter = .Filter & "Enhanced Metafile (*.emf)|*.emf|"
      .Filter = .Filter & "GIF files (*.gif)|*.gif|"
      .Filter = .Filter & "JPEG (*.jfif;*.jpeg;*.jgp)|*.jfif;*.jpeg;*.jpg"
   End With
    
   DialogBox.ShowOpen
   If Not DialogBox.FileName = vbNullString Then
      If Not TGALoader.DrawTGA(ScreenRegionsBox.EditingBox, DialogBox.FileName, ResizeCanvas:=False) Then
         ScreenRegionsBox.EditingBox = LoadPicture(DialogBox.FileName)
      End If
      
      DrawRegions
      ScreenRegionsBox.Show
      ScreenRegionsBox.ZOrder
      ScreenRegionsBox.WindowState = vbNormal
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure displays the load script dialog and gives the command to load the specified script.
Private Sub LoadScriptMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   
   If ScriptCode.NotSaved Then
      If MsgBox("Save the current script first?", vbYesNo Or vbQuestion Or vbDefaultButton1) = vbYes Then SaveCurrentScript
   End If
    
   DialogBox.DialogTitle = "Load Script"
   DialogBox.FileName = vbNullString
   DialogBox.Filter = "SLUDGE Scripts (*.slu)|*.slu"
   DialogBox.ShowOpen
   If Not DialogBox.FileName = vbNullString Then
      StartNewScript
      LoadScript DialogBox.FileName
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure further initializes this window.
Private Sub MDIForm_Activate()
On Error Resume Next
Dim CommandLineArguments As String
   
   Me.WindowState = vbMaximized

   With Settings
      InvertedMenu.Checked = .DisplayInverted
      OtherMenu.Checked = Not .DisplayInverted
      UseSeparateTextLinesMenu.Checked = .UseSeparateTextLines
      UseShiftToAddMenu.Checked = .UseShiftToAdd
      UseShiftToSelectMenu.Checked = .UseShiftToSelect
   End With
   
   PropertiesBox.Show
   ScreenRegionsBox.Show
   ScriptBox.Show
   
   CommandLineArguments = Command$
    
   If CommandLineArguments = vbNullString Then
      ScreenRegionsBox.SetFocus
   ElseIf Not CommandLineArguments = vbNullString Then
      LoadScript CommandLineArguments
      ScriptCode.ManuallyEdited = False
   End If
End Sub

'This procedure initializes this window.
Private Sub MDIForm_Load()
On Error Resume Next
   Me.Width = (Screen.Width / 1.1)
   Me.Height = (Screen.Height / 1.1)
   
   DialogBox.Flags = DialogBox.Flags Or cdlOFNHideReadOnly Xor cdlOFNOverwritePrompt
   ScriptBox.CodeBox.Text = vbNullString
   SaveCurrentScript NewDialogBox:=DialogBox
End Sub

'This procedure gives the command to quit this program after user confirmation.
Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
   If Quit() = vbNo Then
      Cancel = True
      ScriptBoxVisible = True

      PropertiesBox.Show
      ScreenRegionsBox.Show
      ScriptBox.Show
   End If
Endroutine:
   Exit Sub

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to start a new script after user confirmation..
Private Sub NewScriptMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   
   If MsgBox("Start new script?", vbYesNo Or vbQuestion Or vbDefaultButton2) = vbYes Then
      StartNewScript
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure sets screen region the display color.
Private Sub OtherMenu_Click()
On Error GoTo ErrorTrap
   With Settings
      BlockEditing = True
      
      DialogBox.Color = .DisplayColor
      DialogBox.ShowColor
      .DisplayColor = DialogBox.Color
      .DisplayInverted = False
      
      InvertedMenu.Checked = .DisplayInverted
      OtherMenu.Checked = Not .DisplayInverted
      DrawRegions
   End With
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure pastes the text on the clipboard into the script code.
Private Sub PasteMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   
   ScriptBox.CodeBox.SelText = Clipboard.GetText
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure removes the image.
Private Sub RemoveImageMenu_Click()
On Error Resume Next
   BlockEditing = True
   
   ScreenRegionsBox.EditingBox.Picture = Nothing
   
   DrawRegions
   ScreenRegionsBox.Show
   ScreenRegionsBox.ZOrder
   ScreenRegionsBox.WindowState = vbNormal
End Sub

'This procedure gives the command to search for and replace the specified text.
Private Sub ReplaceMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   
   SearchForText FindAndReplace
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to search for and replace the next occurrence of the specified text.
Private Sub ReplaceNextMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   
   SearchForText FindAndReplaceNext
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure displays the save script dialog and gives the command to save the current script.
Private Sub SaveScriptAsMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   
   DialogBox.DialogTitle = "Save Script As"
   DialogBox.FileName = ScriptCode.File
   DialogBox.Filter = "SLUDGE Scripts (*.slu)|*.slu"
   DialogBox.ShowSave
   If Not DialogBox.FileName = vbNullString Then SaveScript DialogBox.FileName
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to save the current script and display a dialog only if the script has no filename.
Private Sub SaveScriptMenu_Click()
On Error GoTo ErrorTrap
   SaveCurrentScript
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure requests the user to specify the screen region handle dimensions.
Private Sub ScreenRegionHandleSizeMenu_Click()
On Error GoTo ErrorTrap
Dim NewHandleSize As Long

   With Settings
      BlockEditing = True
      
      NewHandleSize = CLng(Val(InputBox$("Handle size 1-5 (in pixels.)", , CStr(.HandleSize))))
      If NewHandleSize > 0 And NewHandleSize < 6 Then
         .HandleSize = NewHandleSize
      ElseIf Not NewHandleSize = 0 Then
         MsgBox "The specified handle size is out of range.", vbExclamation
      End If
       
      DrawRegions
   End With
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure selects all script code.
Private Sub SelectAllMenu_Click()
On Error GoTo ErrorTrap
   BlockEditing = True
   
   ScriptBox.CodeBox.SelStart = 0
   ScriptBox.CodeBox.SelLength = Len(ScriptBox.CodeBox.Text)
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure toggles the use separate text lines for screen region definitions option on or off.
Private Sub UseSeparateTextLinesMenu_Click()
On Error GoTo ErrorTrap
   With Settings
      BlockEditing = True
      
      .UseSeparateTextLines = Not .UseSeparateTextLines
      UseSeparateTextLinesMenu.Checked = .UseSeparateTextLines
   End With
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure toggles the use shift key to add option on or off.
Private Sub UseShiftToAddMenu_Click()
On Error GoTo ErrorTrap
   With Settings
      BlockEditing = True
      
      .UseShiftToAdd = Not .UseShiftToAdd
      UseShiftToAddMenu.Checked = .UseShiftToAdd
   End With
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure toggles the use shift key to select option on or off.
Private Sub UseShiftToSelectMenu_Click()
On Error GoTo ErrorTrap
   With Settings
      BlockEditing = True
      
      .UseShiftToSelect = Not .UseShiftToSelect
      UseShiftToSelectMenu.Checked = .UseShiftToSelect
   End With
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure displays the selected window.
Private Sub WindowMenu_Click(Index As Integer)
On Error Resume Next

   BlockEditing = True
    
   If Index = 0 Then
      PropertiesBox.Show
      PropertiesBox.ZOrder
      PropertiesBox.WindowState = vbNormal
   ElseIf Index = 1 Then
      ScreenRegionsBox.Show
      ScreenRegionsBox.ZOrder
      ScreenRegionsBox.WindowState = vbNormal
   ElseIf Index = 2 Then
      ScriptBoxVisible = True
      ScriptBox.Show
      ScriptBox.ZOrder
      ScriptBox.WindowState = vbNormal
   End If
End Sub

