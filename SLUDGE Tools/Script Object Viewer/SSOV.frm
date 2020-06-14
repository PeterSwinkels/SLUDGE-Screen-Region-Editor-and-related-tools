VERSION 5.00
Begin VB.Form SSOVBox 
   Caption         =   "SLUDGE Script Object Viewer"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9165
   ClipControls    =   0   'False
   Icon            =   "SSOV.frx":0000
   ScaleHeight     =   16.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   76.375
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox EventParameterList 
      Height          =   3375
      Left            =   6120
      TabIndex        =   6
      ToolTipText     =   "The events defined for the selected objecttype. Right-click for options."
      Top             =   240
      Width           =   2895
   End
   Begin VB.ListBox ObjectsList 
      Height          =   3375
      Left            =   3120
      TabIndex        =   3
      ToolTipText     =   "The objecttypes and subroutines defined in the script. Right-click for options."
      Top             =   240
      Width           =   2895
   End
   Begin VB.DirListBox DirectoryList 
      Height          =   1440
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "The directories on the selected drive."
      Top             =   720
      Width           =   2895
   End
   Begin VB.DriveListBox DriveList 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "The available drives."
      Top             =   240
      Width           =   2895
   End
   Begin VB.FileListBox FileList 
      Height          =   1455
      Left            =   120
      Pattern         =   "*.SLU"
      TabIndex        =   1
      ToolTipText     =   "The script files in the selected directory. Double-click to open the selected script with the default editor."
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label EventsParametersLabel 
      Caption         =   "Events/Parameters:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label ObjectsLabel 
      Caption         =   "Objects:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.Label PathLabel 
      Caption         =   "Path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.Menu ProgramMenu 
      Caption         =   "&Program"
      Begin VB.Menu CloseProgramMenu 
         Caption         =   "&Close Program"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu SortObjectMenu 
      Caption         =   "Sort &Objects"
      Begin VB.Menu SortByTypeMenu 
         Caption         =   "Sort By &Type"
         Shortcut        =   ^T
      End
      Begin VB.Menu SortByTypeAndNameMenu 
         Caption         =   "Sort By Type and &Name"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu SortEventsParametersMenu 
      Caption         =   "Sort &Events/Parameters"
      Begin VB.Menu SortByNameMenu 
         Caption         =   "Sort By &Name"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu InformationMenu 
      Caption         =   "&Information"
      Begin VB.Menu AboutMenu 
         Caption         =   "&About"
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "SSOVBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's interface.
Option Explicit



'This procedure displays the about dialog box.
Private Sub AboutMenu_Click()
On Error Resume Next
   AboutBox.Show vbModal
End Sub

'This procedure gives the command to quit this program after user confirmation.
Private Sub CloseProgramMenu_Click()
On Error Resume Next
   Unload Me
End Sub

'This procudure changes the current directory.
Private Sub Directorylist_Change()
On Error GoTo ErrorTrap
   FileList.Path = DirectoryList.Path
   ChDir DirectoryList.Path
Endroutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume Endroutine
End Sub

'This procedure changes the current drive.
Private Sub Drivelist_Change()
On Error GoTo ErrorTrap
   DirectoryList.Path = DriveList.Drive
   ChDrive DriveList.Drive
Endroutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume Endroutine
End Sub

'This procedure displays the context menu for the event/parameter list.
Private Sub EventParameterList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   If Button = vbRightButton Then PopupMenu SortEventsParametersMenu
End Sub

'This procedure gives the command to display the objects in the selected script.
Private Sub Filelist_Click()
On Error GoTo ErrorTrap
   ScriptCode ScriptFile:=FileList.List(FileList.ListIndex)
   DisplayObjects GetEventsParameters:=False, EventParameterList:=EventParameterList, ObjectsList:=ObjectsList
Endroutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume Endroutine
End Sub

'This procedure opens the selected script, if the user double clicks it's filename.
Private Sub FileList_DblClick()
On Error GoTo ErrorTrap
   Shell "Explorer.exe " & FileList.List(FileList.ListIndex)
Endroutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume Endroutine
End Sub

'This procedure initializes the interface.
Private Sub Form_Load()
On Error Resume Next
   Me.Width = Screen.Width / 1.5
   Me.Height = Screen.Height / 1.5
End Sub

'This procedure adjusts the size and position of the objects, if this window is resized.
Private Sub Form_Resize()
On Error Resume Next
   DirectoryList.Width = (Me.ScaleWidth - 4) / 3
   DriveList.Width = (Me.ScaleWidth - 4) / 3
   EventParameterList.Width = (Me.ScaleWidth - 4) / 3
   FileList.Width = (Me.ScaleWidth - 4) / 3
   ObjectsList.Width = (Me.ScaleWidth - 4) / 3
   
   DirectoryList.Height = (Me.ScaleHeight - (DriveList.Top + DriveList.Height + 2)) / 2
   EventParameterList.Height = Me.ScaleHeight - 2
   FileList.Height = (Me.ScaleHeight - (DriveList.Top + DriveList.Height + 2)) / 2
   ObjectsList.Height = Me.ScaleHeight - 2
   
   ObjectsLabel.Left = FileList.Left + FileList.Width + 1
   ObjectsList.Left = FileList.Left + FileList.Width + 1
   EventsParametersLabel.Left = ObjectsList.Left + ObjectsList.Width + 1
   EventParameterList.Left = ObjectsList.Left + ObjectsList.Width + 1
   
   FileList.Top = DriveList.Top + DirectoryList.Top + DirectoryList.Height
End Sub

'This procedure gives the command to quit this program after user confirmation.
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   Cancel = (Quit() = vbNo)
End Sub

'This procedure gives the command to display the events/parameters in the selected objecttype.
Private Sub Objectslist_Click()
On Error GoTo ErrorTrap
   DisplayObjects GetEventsParameters:=True, EventParameterList:=EventParameterList, ObjectsList:=ObjectsList
Endroutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume Endroutine
End Sub

'This procedure displays the context menu for the object list.
Private Sub ObjectsList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap
   If Button = vbRightButton Then PopupMenu SortObjectMenu
Endroutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume Endroutine
End Sub

'This procedure toggles between sorting events/parameters by name and no sorting.
Private Sub SortByNameMenu_Click()
On Error GoTo ErrorTrap
   SortByNameMenu.Checked = Not SortByNameMenu.Checked
   DisplayObjects GetEventsParameters:=True, EventParameterList:=EventParameterList, ObjectsList:=ObjectsList
   If SortByNameMenu.Checked Then SortByName EventParameterList
Endroutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume Endroutine
End Sub

'This procedure toggles between sorting objects by type and name, and no sorting.
Private Sub SortByTypeAndNameMenu_Click()
On Error GoTo ErrorTrap
   SortByTypeAndNameMenu.Checked = Not SortByTypeAndNameMenu.Checked
   SortByTypeMenu.Checked = False
   DisplayObjects GetEventsParameters:=False, EventParameterList:=EventParameterList, ObjectsList:=ObjectsList
   If SortByTypeAndNameMenu.Checked Then SortByName ObjectsList
Endroutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume Endroutine
End Sub

'This procedure toggles between sorting objects by type and no sorting.
Private Sub SortByTypeMenu_Click()
On Error GoTo ErrorTrap
   SortByTypeMenu.Checked = Not SortByTypeMenu.Checked
   SortByTypeAndNameMenu.Checked = False
   DisplayObjects GetEventsParameters:=False, EventParameterList:=EventParameterList, ObjectsList:=ObjectsList
   If SortByTypeMenu.Checked Then SortByType ObjectsList
Endroutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume Endroutine
End Sub

