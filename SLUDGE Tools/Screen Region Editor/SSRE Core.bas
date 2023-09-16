Attribute VB_Name = "SSRECore"
'This module contains this program's core procedures.
Option Explicit

'This enumeration lists the corners of a screen region.
Public Enum CornersE
   NoCorner = -1      'No corner.
   UpperLeftCorner    'Upper left corner.
   UpperRightCorner   'Upper right corner.
   LowerRightCorner   'Lower right corner.
   LowerLeftCorner    'Lower left corner.
End Enum

'This enumeration lists the predefined character positions.
Public Enum PredefinedCharacterXYsE
   CharacterXYUserDefined        'The character's position is undefined.
   CharacterXYUpperLeftCorner    'The character's upper left corner position.
   CharacterXYUpperRightCorner   'The character's upper right corner position.
   CharacterXYLowerRightCorner   'The character's lower right corner position.
   CharacterXYLowerLeftCorner    'The character's lower left corner position.
   CharacterXYCenter             'The character's position is in the center of a screen region.
   CharacterXYTopCenter          'The character's position is at the top center of a screen region.
   CharacterXYBottomCenter       'The character's position is at the bottom center of a screen region.
   CharacterXYLeftsideCenter     'The character's position is at the center left side of a screen region.
   CharacterXYRightsideCenter    'The character's position is at the center right side of a screen region.
End Enum

'This enumeration lists the search and replace actions for the script code.
Public Enum SearchActionsE
   Find                  'Find a match.
   FindNext              'Find the next match.
   FindAndReplace        'Find and replace matches.
   FindAndReplaceNext    'Find and replace the next match.
End Enum

'This structure defines the script code.
Public Type ScriptCodeStr
   Code As String              'Contains the script code of the script being edited.
   ManuallyEdited As Boolean   'Indicates whether the script code has been manually edited.
   NotSaved As Boolean         'Indicates whether the script code has been changed since it was last saved.
   File As String              'Contains the name and path of the current script file.
End Type

'This structure defines information about the selected screen region.
Public Type SelectionStr
   ObjectType As String        'The objecttype.
   x1 As Long                  'The horizontal coordinate for the upper left corner.
   y1 As Long                  'The vertical coordinate for the upper left corner.
   x2 As Long                  'The horizontal coordinate for the lower right corner.
   y2 As Long                  'The vertical coordinate for the lower right corner.
   CharacterX As Long          'The horizontal character position.
   CharacterY As Long          'The vertical character position.
   Direction As String         'The character direction.
   Corner As Long              'The selected corner.
   Region As Long              'The selected screen region.
End Type

'This enumeration lists the screen region properties.
Private Enum PropertiesE
   ObjectTypeProperty     'The object type property.
   x1Property             'The upper left corner's horizontal position property.
   y1Property             'The upper left corner's vertical position property.
   x2Property             'The lower right corner's horizontal position property.
   y2Property             'The lower right corner's vertical position property.
   CharacterXProperty     'The character's horizontal position property.
   CharacterYProperty     'The character's vertical position property.
   DirectionProperty      'The character's direction property.
End Enum

'This structure defines the screen region data.
Private Type ScreenRegionsStr
   Properties() As String      'The properties of the screen region.
   CodeOffset() As Long        'The offset of the screen region definition code.
   CodeLength() As Long        'The length of the screen region definition code.
End Type

'This structure defines this program' settings.
Private Type SettingsStr
   DisplayColor As Long                'Defines the color used to display screen regions.
   DisplayInverted As Boolean          'Indicates whether to use inverted colors to display screen regions.
   GridCellHeight As Long              'Defines the height of the screen region editing window grid cells in pixels.
   GridCellWidth As Long               'Defines the width of the screen region editing window grid cells in pixels.
   HandleSize As Long                  'Defines the width of the screen region handles divided by two in pixels.
   UseSeparateTextLines As Boolean     'Indicates whether separate text lines are used for screen region definitions.
   UseShiftToAdd As Boolean            'Indicates whether the shift key needs to be pressed when adding a screen region.
   UseShiftToSelect As Boolean         'Indicates whether the shift key needs to be pressed when selecting a screen region.
End Type

Public BlockEditing As Boolean                          'Indicates whether to temporarily block the editing of the screen regions.
Public EditorHeight As Long                             'Contains the width of the screen region editing window in pixels.
Public EditorWidth As Long                              'Contains the height of the screen region editing window in pixels.
Public LastDirectionUsed As String                      'Contains the direction of most recently added screen region.
Public LastObjectTypeUsed As String                     'Contains the object type of the most recently added screen region.
Public PropertiesBoxVisible As Boolean                  'Indicates whether the properties window is visible.
Public PropertiesChanged As Boolean                     'Indicates whether the properties have been changed.
Public ScreenRegions As ScreenRegionsStr                'Contains the screen regions defined in the script code.
Public ScriptCode As ScriptCodeStr                      'Contains the script code and its properties.
Public ScriptBoxVisible As Boolean                      'Indicates whether the script window is visible.
Public SelectedCharacterXY As PredefinedCharacterXYsE   'Contains the selected predefined character location.
Public Selection As SelectionStr                        'Contains the information about the selected screen region.
Public Settings As SettingsStr                          'Contains this program's settings.

Public Const NO_REGION As Long = -1        'Indicates that no screen region has been selected.

'The Microsoft Windows API constants used.
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002    'Defines the HKEY_LOCAL_MACHINE registry key handle.
Private Const KEY_ALL_ACCESS As Long = &HF003F           'Defines a request for full access to a registry key.
Private Const REG_DWORD As Long = 4                      'Defines the registry value DWORD data type.

'The Microsoft Windows API functions used.
Public Declare Function GetFocus Lib "User32.dll" () As Long
Private Declare Function RegCloseKey Lib "Advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyExA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As String) As Long
'This procedure adds a new screen region using the specified properties.
Public Sub AddRegion(Optional ObjectType As String = vbNullString, Optional x1 As Long = Empty, Optional y1 As Long = Empty, Optional x2 As Long = Empty, Optional y2 As Long = Empty, Optional CharacterX As Long = Empty, Optional CharacterY As Long = Empty, Optional Direction As String = vbNullString)
On Error GoTo ErrorTrap
Dim OffsetShift As Long
Dim Region As Long
Dim RegionDefinition As String

   If Len(ScriptBox.CodeBox.Text) >= ScriptBox.CodeBox.MaxLength Then
      MsgBox "The script code has reached it maximum allowed size of " & CStr(ScriptBox.CodeBox.MaxLength) & " bytes.", vbExclamation
   Else
      Screen.MousePointer = vbHourglass
      With ScreenRegions
         If SafeArrayGetDim(.Properties) = 0 Then
            ReDim .Properties(0 To 0) As String
            ReDim .CodeOffset(0 To 0) As Long
            ReDim .CodeLength(0 To 0) As Long
         End If
         If ObjectType = vbNullString Then ObjectType = "Region"
         If Direction = vbNullString Then Direction = "NULL"
         If x1 > x2 Then Swap x1, x2
         If y1 > y2 Then Swap y1, y2
        
         SetCharacterXY SelectedCharacterXY
         SnapRegionToGrid
         MoveOutsideCodeLine
        
         .Properties(UBound(.Properties())) = ObjectType & "," & x1 & "," & y1 & "," & x2 & "," & y2 & "," & CharacterX & "," & CharacterY & "," & Direction
         RegionDefinition = "addScreenRegion (" & ObjectType & ", " & x1 & ", " & y1 & ", " & x2 & ", " & y2 & ", " & CharacterX & ", " & CharacterY & ", " & Direction & ");"
         .CodeOffset(UBound(.CodeOffset())) = ScriptBox.CodeBox.SelStart
         .CodeLength(UBound(.CodeLength())) = Len(RegionDefinition)
         ScriptBox.CodeBox.SelText = RegionDefinition
         If Settings.UseSeparateTextLines Then ScriptBox.CodeBox.SelText = ScriptBox.CodeBox.SelText & vbCrLf
         ScriptBox.CodeBox.SelStart = .CodeOffset(UBound(.CodeOffset())) + .CodeLength(UBound(.CodeLength())) + 2
       
         OffsetShift = .CodeLength(UBound(.CodeLength())) + 2
         For Region = LBound(.CodeOffset()) To UBound(.CodeOffset()) - 1
            If .CodeOffset(Region) >= .CodeOffset(UBound(.CodeOffset())) Then
               .CodeOffset(Region) = .CodeOffset(Region) + OffsetShift
            End If
         Next Region
        
         ReDim Preserve .Properties(LBound(.Properties()) To UBound(.Properties()) + 1) As String
         ReDim Preserve .CodeOffset(LBound(.CodeOffset()) To UBound(.CodeOffset()) + 1) As Long
         ReDim Preserve .CodeLength(LBound(.CodeLength()) To UBound(.CodeLength()) + 1) As Long
      End With
   End If
Endroutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub
'This procedure returns the path of the folder containing this program's executable (always ending with a backslash.)
Public Function ApplicationPath() As String
On Error GoTo ErrorTrap
Dim Folder As String
   
   Folder = App.Path
   If Not Right$(Folder, 1) = "\" Then Folder = Folder & "\"
   
Endroutine:
   ApplicationPath = Folder
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure changes the specified screen region using the specified properties.
Public Sub ChangeRegion(ChangedRegion As Long, Optional ObjectType As String = vbNullString, Optional x1 As Long = Empty, Optional y1 As Long = Empty, Optional x2 As Long = Empty, Optional y2 As Long = Empty, Optional CharacterX As Long = Empty, Optional CharacterY As Long = Empty, Optional Direction As String = vbNullString)
On Error GoTo ErrorTrap
Dim Cursor As Long
Dim OffsetShift As Long
Dim Region As Long
Dim RegionDefinition As String

   If Not ChangedRegion = NO_REGION Or SafeArrayGetDim(ScreenRegions.Properties) = 0 Then
      If Len(ScriptBox.CodeBox.Text) >= ScriptBox.CodeBox.MaxLength Then
         MsgBox "The script code has reached it maximum allowed size of " & CStr(ScriptBox.CodeBox.MaxLength) & " bytes.", vbExclamation
      Else
         Screen.MousePointer = vbHourglass
         With ScreenRegions
            Cursor = ScriptBox.CodeBox.SelStart
            If x1 > x2 Then Swap x1, x2
            If y1 > y2 Then Swap y1, y2
            SetCharacterXY SelectedCharacterXY
            If Not Direction = vbNullString Then LastDirectionUsed = Direction
            If Not Not ObjectType = vbNullString Then LastObjectTypeUsed = ObjectType
           
            If ObjectType = vbNullString Then GetRegionProperties ChangedRegion, ObjectType
            If x1 = Empty Then GetRegionProperties ChangedRegion, , x1
            If y1 = Empty Then GetRegionProperties ChangedRegion, , , y1
            If x2 = Empty Then GetRegionProperties ChangedRegion, , , , x2
            If y2 = Empty Then GetRegionProperties ChangedRegion, , , , , y2
            If CharacterX = Empty Then GetRegionProperties ChangedRegion, , , , , , CharacterX
            If CharacterY = Empty Then GetRegionProperties ChangedRegion, , , , , , , CharacterY
            If Direction = vbNullString Then GetRegionProperties ChangedRegion, , , , , , , , Direction
           
            SnapRegionToGrid
            
            ScriptBox.CodeBox.SelStart = .CodeOffset(ChangedRegion)
            ScriptBox.CodeBox.SelLength = .CodeLength(ChangedRegion)
            
            .Properties(ChangedRegion) = ObjectType & "," & x1 & "," & y1 & "," & x2 & "," & y2 & "," & CharacterX & "," & CharacterY & "," & Direction
            RegionDefinition = "addScreenRegion (" & ObjectType & ", " & x1 & ", " & y1 & ", " & x2 & ", " & y2 & ", " & CharacterX & ", " & CharacterY & ", " & Direction & ");"
          
            OffsetShift = .CodeLength(ChangedRegion) - Len(RegionDefinition)
            For Region = LBound(.Properties()) To UBound(.Properties()) - 1
               If .CodeOffset(Region) > .CodeOffset(ChangedRegion) Then .CodeOffset(Region) = .CodeOffset(Region) - OffsetShift
            Next Region
            .CodeLength(ChangedRegion) = Len(RegionDefinition)
          
            ScriptBox.CodeBox.SelText = RegionDefinition
            If Cursor > OffsetShift Then
                ScriptBox.CodeBox.SelStart = Cursor - OffsetShift
            ElseIf Not Cursor > OffsetShift Then
                ScriptBox.CodeBox.SelStart = 0
            End If
         End With
      End If
   End If
Endroutine:
Exit Sub
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure checks whether the script code has been manually edited.
Public Sub CheckForManualCodeEditing()
On Error GoTo ErrorTrap
Dim Message As String

   If ScriptCode.ManuallyEdited Then
      ScriptCode.ManuallyEdited = False
      Selection.Region = NO_REGION
     
      Message = "The visually displayed screen regions need to be" & vbCr
      Message = Message & "updated because the script has been manually edited."
      MsgBox Message, vbInformation
     
      If ScriptBox.CodeBox.SelLength = 0 Then
         GetRegionsFromCode ScriptBox.CodeBox.Text, 0
      ElseIf Not ScriptBox.CodeBox.SelLength = 0 Then
         GetRegionsFromCode ScriptBox.CodeBox.SelText, ScriptBox.CodeBox.SelStart
      End If
     
      DrawRegions
      DisplayProperties
      ScreenRegionsBox.ZOrder
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure checks whether the coordinates defined for a screen region are immediates.
Private Function CoordinatesAreImmediates(RegionProperties As String) As Boolean
On Error GoTo ErrorTrap
Dim AreImmediates As Boolean
Dim Position As Long
Dim Properties As String
Dim RegionProperty As PropertiesE
 
   AreImmediates = True
   Properties = RegionProperties
   RegionProperty = ObjectTypeProperty
   Do Until Properties = vbNullString
      Position = InStr(Properties, ",")
      If Position = 0 Then Position = Len(Properties) + 1
      If Not (RegionProperty = ObjectTypeProperty Or RegionProperty = DirectionProperty) Then
         If Not StringIsImmediate(Left$(Properties, Position - 1)) Then
            AreImmediates = False
            Exit Do
         End If
      End If
      RegionProperty = RegionProperty + 1
      Properties = Mid$(Properties, Position + 1)
   Loop
Endroutine:
   CoordinatesAreImmediates = AreImmediates
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure returns the screen region corner selected by the user.
Public Sub DetermineSelectedCorner(x As Long, y As Long)
On Error GoTo ErrorTrap
Dim Region As Long
Dim x1 As Long
Dim x2 As Long
Dim y1 As Long
Dim y2 As Long

   Selection.Corner = NoCorner
   Selection.Region = NO_REGION
   If Not SafeArrayGetDim(ScreenRegions.Properties) = 0 Then
      Screen.MousePointer = vbHourglass

      With ScreenRegions
         For Region = LBound(.Properties()) To UBound(.Properties()) - 1
            GetRegionProperties Region, , x1, y1, x2, y2
            If x >= x1 - Settings.HandleSize And y >= y1 - Settings.HandleSize And x <= x1 + Settings.HandleSize And y <= y1 + Settings.HandleSize Then
               Selection.Region = Region
               Selection.Corner = UpperLeftCorner
            ElseIf x >= x2 - Settings.HandleSize And y >= y1 - Settings.HandleSize And x <= x2 + Settings.HandleSize And y <= y1 + Settings.HandleSize Then
               Selection.Region = Region
               Selection.Corner = UpperRightCorner
            ElseIf x >= x2 - Settings.HandleSize And y >= y2 - Settings.HandleSize And x <= x2 + Settings.HandleSize And y <= y2 + Settings.HandleSize Then
               Selection.Region = Region
               Selection.Corner = LowerRightCorner
            ElseIf x >= x1 - Settings.HandleSize And y >= y2 - Settings.HandleSize And x <= x1 + Settings.HandleSize And y <= y2 + Settings.HandleSize Then
               Selection.Region = Region
               Selection.Corner = LowerLeftCorner
            End If
         Next Region
      End With
   End If
Endroutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure returns the screen region selected by the user.
Public Sub DetermineSelectedRegion(x As Long, y As Long)
On Error GoTo ErrorTrap
Dim Region As Long
Dim x1 As Long
Dim x2 As Long
Dim y1 As Long
Dim y2 As Long

   Selection.Corner = NoCorner
   Selection.Region = NO_REGION
   If Not SafeArrayGetDim(ScreenRegions.Properties) = 0 Then
      Screen.MousePointer = vbHourglass
      For Region = LBound(ScreenRegions.Properties()) To UBound(ScreenRegions.Properties()) - 1
         GetRegionProperties Region, , x1, y1, x2, y2
         If x >= x1 And x <= x2 And y >= y1 And y <= y2 Then Selection.Region = Region
      Next Region
   End If
Endroutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure displays the selected screen region's properties.
Public Sub DisplayProperties()
On Error GoTo ErrorTrap
   If PropertiesBoxVisible Then
      With PropertiesBox
         .ObjectTypeBox.Text = Selection.ObjectType
         .x1Box.Text = CStr(Selection.x1)
         .y1Box.Text = CStr(Selection.y1)
         .x2Box.Text = CStr(Selection.x2)
         .y2Box.Text = CStr(Selection.y2)
         .CharacterXBox.Text = CStr(Selection.CharacterX)
         .CharacterYBox.Text = CStr(Selection.CharacterY)
         .DirectionBox.Text = Selection.Direction
         If Selection.Region = NO_REGION Then
            .CharacterXYBox.ListIndex = CharacterXYUserDefined
            .RankBox.Text = vbNullString
         ElseIf Not Selection.Region = NO_REGION Then
            .CharacterXYBox.ListIndex = GetCharacterXY()
            .RankBox.Text = Selection.Region + 1 & " of " & UBound(ScreenRegions.Properties())
         End If
      End With
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure draws the selected screen regions.
Public Sub DrawRegions()
On Error GoTo ErrorTrap
Dim CharacterX As Long
Dim CharacterY As Long
Dim Direction As String
Dim Region As Long
Dim x1 As Long
Dim x2 As Long
Dim y1 As Long
Dim y2 As Long

   ScreenRegionsBox.EditingBox.Cls
   If Not SafeArrayGetDim(ScreenRegions.Properties) = 0 Then
      Screen.MousePointer = vbHourglass
   
      If Settings.DisplayInverted Then
         ScreenRegionsBox.EditingBox.DrawMode = vbInvert
      ElseIf Not Settings.DisplayInverted Then
         ScreenRegionsBox.EditingBox.DrawMode = vbCopyPen
         ScreenRegionsBox.EditingBox.ForeColor = Settings.DisplayColor
      End If
      With ScreenRegions
         For Region = LBound(.Properties()) To UBound(.Properties()) - 1
            GetRegionProperties Region, , x1, y1, x2, y2, CharacterX, CharacterY, Direction
            If x1 = x2 And y1 = y2 Then
               ScreenRegionsBox.EditingBox.PSet (x1, y1)
            ElseIf x1 = x2 Xor y1 = y2 Then
               ScreenRegionsBox.EditingBox.Line (x1, y1)-(x2, y2)
            Else
               ScreenRegionsBox.EditingBox.Line (x1, y1)-(x2, y2), , B
            End If
            
            ScreenRegionsBox.EditingBox.Circle (CharacterX, CharacterY), 5
            If Direction = "NORTH" Then
               ScreenRegionsBox.EditingBox.Line (CharacterX, CharacterY - 5)-(CharacterX, CharacterY - 15)
            ElseIf Direction = "NORTHEAST" Then
               ScreenRegionsBox.EditingBox.Line (CharacterX + 5, CharacterY - 5)-(CharacterX + 15, CharacterY - 15)
            ElseIf Direction = "EAST" Then
               ScreenRegionsBox.EditingBox.Line (CharacterX + 5, CharacterY)-(CharacterX + 15, CharacterY)
            ElseIf Direction = "SOUTHEAST" Then
               ScreenRegionsBox.EditingBox.Line (CharacterX + 5, CharacterY + 5)-(CharacterX + 15, CharacterY + 15)
            ElseIf Direction = "SOUTH" Then
               ScreenRegionsBox.EditingBox.Line (CharacterX, CharacterY + 5)-(CharacterX, CharacterY + 15)
            ElseIf Direction = "SOUTHWEST" Then
               ScreenRegionsBox.EditingBox.Line (CharacterX - 5, CharacterY + 5)-(CharacterX - 15, CharacterY + 15)
            ElseIf Direction = "WEST" Then
               ScreenRegionsBox.EditingBox.Line (CharacterX - 5, CharacterY)-(CharacterX - 15, CharacterY)
            ElseIf Direction = "NORTHWEST" Then
               ScreenRegionsBox.EditingBox.Line (CharacterX - 5, CharacterY - 5)-(CharacterX - 15, CharacterY - 15)
            End If
            ScreenRegionsBox.EditingBox.Circle Step(0, 0), 1
                      
            ScreenRegionsBox.EditingBox.FillColor = Settings.DisplayColor
            If Region = Selection.Region Then
               ScreenRegionsBox.EditingBox.FillStyle = vbFSTransparent
               ScreenRegionsBox.EditingBox.Line (x1 - Settings.HandleSize, y1 - Settings.HandleSize)-Step(Settings.HandleSize * 2, Settings.HandleSize * 2), , B
               ScreenRegionsBox.EditingBox.Line (x2 - Settings.HandleSize, y1 - Settings.HandleSize)-Step(Settings.HandleSize * 2, Settings.HandleSize * 2), , B
               ScreenRegionsBox.EditingBox.Line (x2 - Settings.HandleSize, y2 - Settings.HandleSize)-Step(Settings.HandleSize * 2, Settings.HandleSize * 2), , B
               ScreenRegionsBox.EditingBox.Line (x1 - Settings.HandleSize, y2 - Settings.HandleSize)-Step(Settings.HandleSize * 2, Settings.HandleSize * 2), , B
               
               ScreenRegionsBox.EditingBox.FillStyle = vbFSSolid
               If Selection.Corner = UpperLeftCorner Then ScreenRegionsBox.EditingBox.Line (x1 - Settings.HandleSize, y1 - Settings.HandleSize)-Step(Settings.HandleSize * 2, Settings.HandleSize * 2), , B
               If Selection.Corner = UpperRightCorner Then ScreenRegionsBox.EditingBox.Line (x2 - Settings.HandleSize, y1 - Settings.HandleSize)-Step(Settings.HandleSize * 2, Settings.HandleSize * 2), , B
               If Selection.Corner = LowerRightCorner Then ScreenRegionsBox.EditingBox.Line (x2 - Settings.HandleSize, y2 - Settings.HandleSize)-Step(Settings.HandleSize * 2, Settings.HandleSize * 2), , B
               If Selection.Corner = LowerLeftCorner Then ScreenRegionsBox.EditingBox.Line (x1 - Settings.HandleSize, y2 - Settings.HandleSize)-Step(Settings.HandleSize * 2, Settings.HandleSize * 2), , B
            End If
            ScreenRegionsBox.EditingBox.FillStyle = vbFSTransparent
         Next Region
      End With
      ScreenRegionsBox.EditingBox.DrawMode = vbInvert
   End If
Endroutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure returns the predefined value used for the character x/y properties.
Private Function GetCharacterXY() As Long
On Error GoTo ErrorTrap
Dim CenterX As Long
Dim CenterY As Long
Dim CharacterXY As PredefinedCharacterXYsE

   With Selection
      CenterX = .x1 + ((.x2 - .x1) / 2)
      CenterY = .y1 + ((.y2 - .y1) / 2)
    
      If .CharacterX = .x1 And .CharacterY = .y1 Then
         CharacterXY = CharacterXYUpperLeftCorner
      ElseIf .CharacterX = .x2 And .CharacterY = .y1 Then
         CharacterXY = CharacterXYUpperRightCorner
      ElseIf .CharacterX = .x2 And .CharacterY = .y2 Then
         CharacterXY = CharacterXYLowerRightCorner
      ElseIf .CharacterX = .x1 And .CharacterY = .y2 Then
         CharacterXY = CharacterXYLowerLeftCorner
      ElseIf .CharacterX = CenterX And .CharacterY = CenterY Then
         CharacterXY = CharacterXYCenter
      ElseIf .CharacterX = CenterX And .CharacterY = .y1 Then
         CharacterXY = CharacterXYTopCenter
      ElseIf .CharacterX = CenterX And .CharacterY = .y2 Then
         CharacterXY = CharacterXYBottomCenter
      ElseIf .CharacterX = .x1 And .CharacterY = CenterY Then
         CharacterXY = CharacterXYLeftsideCenter
      ElseIf .CharacterX = .x2 And .CharacterY = CenterY Then
         CharacterXY = CharacterXYRightsideCenter
      Else
         CharacterXY = CharacterXYUserDefined
      End If
   End With
Endroutine:
   GetCharacterXY = CharacterXY
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure gets the properties for the specified screen region and returns the resulting values.
Public Sub GetRegionProperties(ReferedRegion As Long, Optional ObjectType As String = vbNullString, Optional x1 As Long = Empty, Optional y1 As Long = Empty, Optional x2 As Long = Empty, Optional y2 As Long = Empty, Optional CharacterX As Long = Empty, Optional CharacterY As Long = Empty, Optional Direction As String = vbNullString)
On Error GoTo ErrorTrap
Dim Position As Long
Dim RegionProperties As String
Dim RegionProperty As PropertiesE

   If ReferedRegion = NO_REGION Or SafeArrayGetDim(ScreenRegions.Properties) = 0 Then
      RegionProperties = String$(8, ",")
   ElseIf Not (ReferedRegion = NO_REGION Or SafeArrayGetDim(ScreenRegions.Properties) = 0) Then
      RegionProperties = ScreenRegions.Properties(ReferedRegion)
   End If
   
   RegionProperty = ObjectTypeProperty
   Do Until RegionProperties = vbNullString
      Position = InStr(RegionProperties, ",")
      If Position = 0 Then Position = Len(RegionProperties) + 1
      If RegionProperty = ObjectTypeProperty Then ObjectType = Left$(RegionProperties, Position - 1)
      If RegionProperty = x1Property Then x1 = CLng(Val(Left$(RegionProperties, Position - 1)))
      If RegionProperty = y1Property Then y1 = CLng(Val(Left$(RegionProperties, Position - 1)))
      If RegionProperty = x2Property Then x2 = CLng(Val(Left$(RegionProperties, Position - 1)))
      If RegionProperty = y2Property Then y2 = CLng(Val(Left$(RegionProperties, Position - 1)))
      If RegionProperty = CharacterXProperty Then CharacterX = CLng(Val(Left$(RegionProperties, Position - 1)))
      If RegionProperty = CharacterYProperty Then CharacterY = CLng(Val(Left$(RegionProperties, Position - 1)))
      If RegionProperty = DirectionProperty Then Direction = Left$(RegionProperties, Position - 1)
      RegionProperty = RegionProperty + 1
      RegionProperties = Mid$(RegionProperties, Position + 1)
   Loop
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure scans the specified selection of the script code for screen region definitions.
Public Sub GetRegionsFromCode(SelectedCode As String, SelectionOffset As Long)
On Error GoTo ErrorTrap
Dim Character As String
Dim Code As String
Dim CodeLine As String
Dim EndOfCodeLine As Long
Dim InComment As Boolean
Dim InFileName As Boolean
Dim InString As Boolean
Dim Offset As Long
Dim Position As Long
Dim PreviousCharacter As String
Dim RegionProperties As String

   Screen.MousePointer = vbHourglass
   Code = SelectedCode
   With ScreenRegions
      ReDim .Properties(0 To 0) As String
      ReDim .CodeOffset(0 To 0) As Long
      ReDim .CodeLength(0 To 0) As Long
   
      Code = Replace(Code, vbLf, " ")
      Code = Replace(Code, vbTab, " ")
      InComment = False
      InFileName = False
      InString = False
      Offset = 0
      For Position = 1 To Len(Code)
         PreviousCharacter = Character
         Character = Mid$(Code, Position, 1)
       
         If InComment Then
            If Character = vbCr Then InComment = False
         ElseIf InFileName Then
            If Character = "'" Then InFileName = False
         ElseIf InString Then
            If Character = """" And Not PreviousCharacter = "\" Then InString = False
         Else
            If Character = "#" Then InComment = True
            If Character = "'" Then InFileName = True
            If Character = """" Then InString = True
         End If
   
         If Not (InComment Or InFileName Or InString) Then
            If Offset = 0 Then
               If Mid$(Code, Position, Len("addScreenRegion")) = "addScreenRegion" Then
                  Offset = Position
               End If
            ElseIf Offset > 0 Then
               If Character = ";" Or Character = "}" Then
                  CodeLine = Mid$(Code, Offset, Position - Offset + 1)
                  RegionProperties = Mid$(Replace(CodeLine, " ", vbNullString), Len("addScreenRegion(") + 1)
                  EndOfCodeLine = InStr(RegionProperties, ")")
                  If EndOfCodeLine > 0 Then
                     RegionProperties = Left$(RegionProperties, EndOfCodeLine - 1)
                     If CoordinatesAreImmediates(RegionProperties) Then
                        .CodeOffset(UBound(.CodeOffset())) = (Offset - 1) + SelectionOffset
                        .CodeLength(UBound(.CodeLength())) = Len(CodeLine)
                        .Properties(UBound(.Properties())) = RegionProperties
                        ReDim Preserve .Properties(LBound(.Properties()) To UBound(.Properties()) + 1) As String
                        ReDim Preserve .CodeOffset(LBound(.CodeOffset()) To UBound(.CodeOffset()) + 1) As Long
                        ReDim Preserve .CodeLength(LBound(.CodeLength()) To UBound(.CodeLength()) + 1) As Long
                     End If
                  End If
                  Offset = 0
               End If
            End If
         End If
      Next Position
   End With
Endroutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure handles any errors that occur.
Public Function HandleError(Optional ReturnPreviousChoice As Boolean = True) As Long
Dim ErrorCode As Long
Dim Message As String
Static Choice As Long

   If Not ReturnPreviousChoice Then
      ErrorCode = Err.Number
      Message = Err.Description
      
      On Error Resume Next
      Message = Message & "." & vbCr
      Message = Message & "Error code: " & ErrorCode
      Choice = MsgBox(Message, vbAbortRetryIgnore Or vbExclamation Or vbDefaultButton2)
      If Choice = vbAbort Then End
      If Choice = vbIgnore Then Screen.MousePointer = vbDefault
   End If

   HandleError = Choice
End Function

'This procedure initializes this program.
Private Sub InitializeProgram()
On Error GoTo ErrorTrap
Dim KeyHandle As Long

   With Settings
      .DisplayColor = vbWhite
      .DisplayInverted = True
      .GridCellHeight = 1
      .GridCellWidth = 1
      .HandleSize = 2
      .UseSeparateTextLines = True
      .UseShiftToAdd = False
      .UseShiftToSelect = False
   End With

   RegOpenKeyExA HKEY_LOCAL_MACHINE, "SOFTWARE\Hungry Software\SLUDGE Compiler", CLng(0), KEY_ALL_ACCESS, KeyHandle
   RegQueryValueExA KeyHandle, "utilityWidth", CLng(0), REG_DWORD, EditorWidth, Len(EditorWidth)
   RegQueryValueExA KeyHandle, "utilityHeight", CLng(0), REG_DWORD, EditorHeight, Len(EditorHeight)
   RegCloseKey KeyHandle
   
   If EditorWidth = 0 Then EditorWidth = 640
   If EditorHeight = 0 Then EditorHeight = 480
   
   InitializeScript
   
   DrawRegions
   ScriptBox.CodeBox = vbNullString
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure initializes the script.
Private Sub InitializeScript()
On Error GoTo ErrorTrap
   BlockEditing = False
   LastDirectionUsed = "NULL"
   LastObjectTypeUsed = vbNullString
   PropertiesBoxVisible = False
   PropertiesChanged = False
   ScriptBoxVisible = True
   SelectedCharacterXY = CharacterXYUserDefined
   
   Erase ScreenRegions.Properties(), ScreenRegions.CodeOffset(), ScreenRegions.CodeLength()

   With ScriptCode
      .Code = vbNullString
      .File = vbNullString
      .ManuallyEdited = False
      .NotSaved = False
   End With

   With Selection
      .ObjectType = vbNullString
      .x1 = 0
      .y1 = 0
      .x2 = 0
      .y2 = 0
      .CharacterX = 0
      .CharacterY = 0
      .Direction = "NULL"
      .Corner = NoCorner
      .Region = NO_REGION
   End With
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub


'This procedure loads the specified script file.
Public Sub LoadScript(FileName As String)
On Error GoTo ErrorTrap
Dim Code As String
Dim CodeLine As String
Dim FileHandle As Long

   If Left$(FileName, 1) = """" Then FileName = Mid$(FileName, 2)
   If Right$(FileName, 1) = """" Then FileName = Left$(FileName, Len(FileName) - 1)
   
   If FileLen(FileName) <= ScriptBox.CodeBox.MaxLength Then
      Screen.MousePointer = vbHourglass
      Code = vbNullString
      FileHandle = FreeFile()
      FileName = RemoveQuotes(FileName)
      Open FileName For Input Lock Read Write As FileHandle
         Do Until EOF(FileHandle)
            Line Input #FileHandle, CodeLine
            Code = Code & CodeLine & vbCrLf
         Loop
      Close FileHandle
      
      With ScriptCode
         .NotSaved = False
         .File = FileName
         ScriptBox.Caption = .File
      End With

      ScriptBox.CodeBox.Text = Code
      ScriptBox.SetFocus
      ScriptBox.ZOrder
   Else
      MsgBox "This program cannot load" & vbCr & "scripts larger than " & CStr(ScriptBox.CodeBox.MaxLength) & " bytes.", vbInformation
   End If
Endroutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure loads this program's settings.
Private Sub LoadSettings()
On Error GoTo ErrorTrap
Dim Data As String
Dim FileHandle As Long
Dim ValueData As String
Dim ValueName As String

   If Not Dir$(ApplicationPath & App.Title & ".ini") = vbNullString Then
      FileHandle = FreeFile()
      Open ApplicationPath & App.Title & ".ini" For Input Lock Read Write As FileHandle
         Do Until EOF(FileHandle)
            Line Input #FileHandle, Data
            Data = Trim$(Data)
            If InStr(Data, "=") > 0 Then
               ValueName = UCase$(Trim$(Left$(Data, InStr(Data, "=") - 1)))
               ValueData = UCase$(Trim$(Mid$(Data, Len(ValueName) + 2)))
               
               On Error Resume Next
               With Settings
                  If ValueName = "CELLHEIGHT" Then .GridCellHeight = CLng(Val(ValueData))
                  If ValueName = "CELLWIDTH" Then .GridCellWidth = CLng(Val(ValueData))
                  If ValueName = "DISPLAYCOLOR" Then .DisplayColor = CLng(Val("&H" & ValueData & "&"))
                  If ValueName = "DISPLAYINVERTED" Then .DisplayInverted = CBool(ValueData)
                  If ValueName = "HANDLESIZE" Then .HandleSize = CLng(Val(ValueData))
                  If ValueName = "SEPARATELINES" Then .UseSeparateTextLines = CBool(ValueData)
                  If ValueName = "USESHIFTOADD" Then .UseShiftToAdd = CBool(ValueData)
                  If ValueName = "USESHIFTOSELECT" Then .UseShiftToSelect = CBool(ValueData)
               End With
               On Error GoTo ErrorTrap
            End If
         Loop
      Close FileHandle
   End If
Endroutine:
   With Settings
      If Not (.DisplayColor >= vbBlack And .DisplayColor <= vbWhite) Then .DisplayColor = vbWhite
      If Not (.GridCellHeight > 0 And .GridCellHeight < 101) Then .GridCellHeight = 1
      If Not (.GridCellWidth > 0 And .GridCellWidth < 101) Then .GridCellWidth = 1
      If Not (.HandleSize > 0 And .HandleSize < 6) Then .HandleSize = 2
   End With
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure initializes this program.
Private Sub Main()
On Error GoTo ErrorTrap
   InitializeProgram
   LoadSettings
  
   SSREBox.Show
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure moves the script code box cursor to the end of the nearest code line if it is at the middle of a code line.
Private Sub MoveOutsideCodeLine()
On Error GoTo ErrorTrap
Dim Character As String
Dim Code As String
Dim Cursor As Long
Dim InComment As Boolean
Dim InFileName As Boolean
Dim InString As Boolean
Dim Offset As Long
Dim Position As Long
Dim PreviousCharacter As String
Dim WasInComment As Boolean

   Screen.MousePointer = vbHourglass
   Cursor = ScriptBox.CodeBox.SelStart
   Code = ScriptCode.Code
   Code = Replace(Code, vbLf, " ")
   Code = Replace(Code, vbTab, " ")
   InComment = False
   InFileName = False
   InString = False
   Offset = 0
   WasInComment = False
   For Position = 1 To Len(Code)
      PreviousCharacter = Character
      Character = Mid$(Code, Position, 1)
      
      If InComment Then
         If Character = vbCr Then
            InComment = False
            WasInComment = True
         End If
      ElseIf InFileName Then
         If Character = "'" Then InFileName = False
      ElseIf InString Then
         If Character = """" And Not PreviousCharacter = "\" Then InString = False
      Else
         If Character = "#" Then
            InComment = True
            WasInComment = False
         End If
         If Character = "'" Then InFileName = True
         If Character = """" Then InString = True
      End If
      
      If Not (InFileName Or InString) Then
         If Offset = 0 Then
            If ((Character = "{" Or (UCase$(Character) >= "A" And UCase$(Character) <= "Z") Or Character = "}") And Not InComment) Or (Character = "#" And InComment) Then
               Offset = Position
            End If
         End If
         If Offset > 0 Then
            If ((Character = "{" Or Character = ";" Or Character = "}") And Not InComment) Or (Character = vbCr And WasInComment) Then
               If Cursor >= Offset And Cursor + 1 <= Position Then
                  Cursor = Position
                  If Mid$(ScriptCode.Code, Cursor + 1, 1) = vbLf Then Cursor = Cursor - 1
                  Exit For
               End If
               Offset = 0
            End If
         End If
      End If
   Next Position
   ScriptBox.CodeBox.SelStart = Cursor
Endroutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure removes, if present, the quotes enclosing the specified path.
Private Function RemoveQuotes(Path As String) As String
On Error GoTo ErrorTrap
Dim ModifiedPath As String

   ModifiedPath = Path
   If Left$(ModifiedPath, 1) = """" Then ModifiedPath = Mid$(ModifiedPath, 2)
   If Right$(ModifiedPath, 1) = """" Then ModifiedPath = Left$(ModifiedPath, Len(ModifiedPath) - 1)
Endroutine:
   RemoveQuotes = ModifiedPath
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure quits this program after user confirmation.
Public Function Quit() As Long
On Error GoTo ErrorTrap
Dim Choice As Long

   Choice = MsgBox("Close the script and quit?", vbYesNo Or vbQuestion Or vbDefaultButton2)

   If Choice = vbYes Then
      SaveSettings
      Unload SSREBox
   End If
Endroutine:
   Quit = Choice
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure removes the specified screen region.
Public Sub RemoveRegion(RemovedRegion As Long)
On Error GoTo ErrorTrap
Dim CodeOffset As Long
Dim Cursor As Long
Dim Index As Long
Dim OffsetShift As Long

   If Not RemovedRegion = NO_REGION Or SafeArrayGetDim(ScreenRegions.Properties) = 0 Then
      Screen.MousePointer = vbHourglass
      With ScreenRegions
         Cursor = ScriptBox.CodeBox.SelStart
         ScriptBox.CodeBox.SelStart = .CodeOffset(RemovedRegion)
         ScriptBox.CodeBox.SelLength = .CodeLength(RemovedRegion)
         ScriptBox.CodeBox.SelText = vbNullString
          
         CodeOffset = .CodeOffset(RemovedRegion)
         OffsetShift = .CodeLength(RemovedRegion)
       
         If Settings.UseSeparateTextLines Then
            If Mid$(ScriptBox.CodeBox.Text, ScriptBox.CodeBox.SelStart + 1, 2) = vbCrLf Then
               ScriptBox.CodeBox.SelLength = 2
               ScriptBox.CodeBox.SelText = vbNullString
               OffsetShift = OffsetShift + 2
            ElseIf Mid$(ScriptBox.CodeBox.Text, Cursor - 1, 1) = vbCr Then
               ScriptBox.CodeBox.SelLength = 1
               ScriptBox.CodeBox.SelText = vbNullString
               OffsetShift = OffsetShift + 1
            End If
         End If
        
         For Index = LBound(.Properties()) To UBound(.Properties()) - 1
            If Index >= RemovedRegion Then
               .Properties(Index) = .Properties(Index + 1)
               .CodeOffset(Index) = .CodeOffset(Index + 1)
               If .CodeOffset(Index + 1) > CodeOffset Then .CodeOffset(Index) = .CodeOffset(Index + 1) - OffsetShift
               .CodeLength(Index) = .CodeLength(Index + 1)
            ElseIf Index < RemovedRegion Then
               .CodeOffset(Index) = .CodeOffset(Index)
               If .CodeOffset(Index) > CodeOffset Then .CodeOffset(Index) = .CodeOffset(Index) - OffsetShift
            End If
         Next Index
        
         ReDim Preserve .Properties(LBound(.Properties()) To UBound(.Properties()) - 1) As String
         ReDim Preserve .CodeOffset(LBound(.CodeOffset()) To UBound(.CodeOffset()) - 1) As Long
         ReDim Preserve .CodeLength(LBound(.CodeLength()) To UBound(.CodeLength()) - 1) As Long
         If Cursor > OffsetShift Then
             ScriptBox.CodeBox.SelStart = Cursor - OffsetShift
         ElseIf Not Cursor > OffsetShift Then
             ScriptBox.CodeBox.SelStart = 0
         End If
      End With
   End If
Endroutine:
   Screen.MousePointer = vbDefault
Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to save the current script.
Public Sub SaveCurrentScript(Optional NewDialogBox As CommonDialog = Nothing)
On Error GoTo ErrorTrap
Static CurrentDialogBox As CommonDialog

   BlockEditing = True
   If NewDialogBox Is Nothing Then
      With CurrentDialogBox
         .DialogTitle = "Save Script"
         .FileName = ScriptCode.File
         .Filter = "SLUDGE Scripts (*.slu)|*.slu"
         If ScriptCode.File = vbNullString Then .ShowSave
         If Not .FileName = vbNullString Then SaveScript .FileName
      End With
   Else
      Set CurrentDialogBox = NewDialogBox
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub



'This procedure saves the specified script file.
Public Sub SaveScript(FileName As String)
On Error GoTo ErrorTrap
Dim FileHandle As Long
   
   Screen.MousePointer = vbHourglass
   If Left$(FileName, 1) = """" Then FileName = Mid$(FileName, 2)
   If Right$(FileName, 1) = """" Then FileName = Left$(FileName, Len(FileName) - 1)
   
   FileHandle = FreeFile()
   FileName = RemoveQuotes(FileName)
   With ScriptCode
      Open FileName For Output Lock Read Write As FileHandle
         Print #FileHandle, .Code;
      Close FileHandle
      
      .NotSaved = False
      .File = FileName
      ScriptBox.Caption = .File
   End With
Endroutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure saves this program's settings.
Private Sub SaveSettings()
On Error GoTo ErrorTrap
Dim FileHandle As Long

   FileHandle = FreeFile()
   Open ApplicationPath & App.Title & ".ini" For Output Lock Read Write As FileHandle
      With Settings
         Print #FileHandle, "[GRID]"
         Print #FileHandle, "CELLHEIGHT="; CStr(.GridCellHeight)
         Print #FileHandle, "CELLWIDTH="; CStr(.GridCellWidth)
         Print #FileHandle,
         Print #FileHandle, "[EDITOR]"
         Print #FileHandle, "DISPLAYCOLOR="; Hex$(.DisplayColor)
         Print #FileHandle, "DISPLAYINVERTED="; CStr(.DisplayInverted)
         Print #FileHandle, "HANDLESIZE="; CStr(.HandleSize)
         Print #FileHandle, "SEPARATELINES="; CStr(.UseSeparateTextLines)
         Print #FileHandle, "USESHIFTOADD="; CStr(.UseShiftToAdd)
         Print #FileHandle, "USESHIFTOSELECT="; CStr(.UseShiftToSelect)
      End With
   Close FileHandle
   
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure searches for the specified text and replaces it with the specified text, if requested.
Public Sub SearchForText(Action As Long)
On Error GoTo ErrorTrap
Dim Position As Long
Static SearchTextReplacement As String
Static SearchText As String

   If Action = Find Or (Action = FindNext And SearchText = vbNullString) Then
      SearchText = InputBox$("Enter search text.", , SearchText)
   ElseIf Action = FindAndReplace Or (Action = FindAndReplaceNext And SearchText = vbNullString) Then
      SearchText = InputBox$("Enter the text to replace.", , SearchText)
   End If
    
   If Action = FindAndReplace Or (Action = FindAndReplaceNext And SearchTextReplacement = vbNullString) Then
      SearchTextReplacement = InputBox$("Enter the text to replace """ & SearchText & """ with.", , SearchTextReplacement)
   End If
    
   If Not SearchText = vbNullString Then
      With ScriptBox.CodeBox
         Position = InStr(.SelStart + 2, .Text, SearchText, vbTextCompare)
         If Position = 0 Then Position = InStr(1, .Text, SearchText, vbTextCompare)
        
         If Position = 0 Then
            MsgBox "Could not find the specified text.", vbInformation
         ElseIf Position > 0 Then
            .SelStart = Position - 1
            If Action = FindAndReplace Or Action = FindAndReplaceNext Then
               .SelLength = Len(SearchText)
               .SelText = SearchTextReplacement
               If Action = FindAndReplaceNext Then .SelStart = .SelStart + Len(SearchTextReplacement)
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

'This procedure changes the character x/y properties to a predefined value.
Public Sub SetCharacterXY(CharacterXY As PredefinedCharacterXYsE)
On Error GoTo ErrorTrap
Dim CenterX As String
Dim CenterY As String

   With Selection
      CenterX = .x1 + ((.x2 - .x1) / 2)
      CenterY = .y1 + ((.y2 - .y1) / 2)
    
      If CharacterXY = CharacterXYUpperLeftCorner Then
         .CharacterX = .x1
         .CharacterY = .y1
      ElseIf CharacterXY = CharacterXYUpperRightCorner Then
         .CharacterX = .x2
         .CharacterY = .y1
      ElseIf CharacterXY = CharacterXYLowerRightCorner Then
         .CharacterX = .x2
         .CharacterY = .y2
      ElseIf CharacterXY = CharacterXYLowerLeftCorner Then
         .CharacterX = .x1
         .CharacterY = .y2
      ElseIf CharacterXY = CharacterXYCenter Then
         .CharacterX = CenterX
         .CharacterY = CenterY
      ElseIf CharacterXY = CharacterXYTopCenter Then
         .CharacterX = CenterX
         .CharacterY = .y1
      ElseIf CharacterXY = CharacterXYBottomCenter Then
         .CharacterX = CenterX
         .CharacterY = .y2
      ElseIf CharacterXY = CharacterXYLeftsideCenter Then
         .CharacterX = .x1
         .CharacterY = CenterY
      ElseIf CharacterXY = CharacterXYRightsideCenter Then
         .CharacterX = .x2
         .CharacterY = CenterY
      End If
   End With
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure adjusts the selected screen region's coordinates to the grid cell dimensions.
Private Sub SnapRegionToGrid()
On Error GoTo ErrorTrap
   With Selection
      .x1 = CInt(.x1 / Settings.GridCellWidth) * Settings.GridCellWidth
      .y1 = CInt(.y1 / Settings.GridCellHeight) * Settings.GridCellHeight
      .x2 = CInt(.x2 / Settings.GridCellWidth) * Settings.GridCellWidth
      .y2 = CInt(.y2 / Settings.GridCellHeight) * Settings.GridCellHeight
      .CharacterX = CInt(.CharacterX / Settings.GridCellWidth) * Settings.GridCellWidth
      .CharacterY = CInt(.CharacterY / Settings.GridCellHeight) * Settings.GridCellHeight
   End With
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure starts a new script.
Public Sub StartNewScript()
On Error GoTo ErrorTrap
   Unload PropertiesBox
   Unload ScreenRegionsBox
   Unload ScriptBox
   
   InitializeScript
   
   PropertiesBox.Show
   ScreenRegionsBox.Show
   ScriptBox.CodeBox = vbNullString
   ScriptBox.Show
   ScreenRegionsBox.SetFocus
   
   DrawRegions
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure checks if the specified string contains an immediate value.
Private Function StringIsImmediate(Number As String) As Boolean
On Error GoTo ErrorTrap
Dim IsImmediate As Boolean
Dim Position As Long
   
   IsImmediate = True
   For Position = 1 To Len(Number)
      If InStr("0123456789", Mid$(Number, Position, 1)) = 0 Then
         IsImmediate = False
         Exit For
      End If
   Next Position
Endroutine:
   StringIsImmediate = IsImmediate
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure exchanges the two specified values with each other.
Private Sub Swap(Variable1 As Variant, Variable2 As Variant)
On Error GoTo ErrorTrap
Dim Variable3 As Variant

   Variable3 = Variable1
   Variable1 = Variable2
   Variable2 = Variable3
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure updates the selected screen region with properties specified in the properties window.
Public Sub UpdateRegion()
On Error GoTo ErrorTrap
   If Not (ScriptCode.ManuallyEdited Or (Not PropertiesBoxVisible) Or Selection.Region = NO_REGION) Then
      With PropertiesBox
         ChangeRegion Selection.Region, .ObjectTypeBox.Text, CLng(Val(.x1Box.Text)), CLng(Val(.y1Box.Text)), CLng(Val(.x2Box.Text)), CLng(Val(.y2Box.Text)), CLng(Val(.CharacterXBox.Text)), CLng(Val(.CharacterYBox.Text)), .DirectionBox.Text
      End With
      
      With Selection
         GetRegionProperties .Region, .ObjectType, .x1, .y1, .x2, .y2, .CharacterX, .CharacterY, .Direction
      End With
      
      DrawRegions
   End If
Endroutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume Endroutine
   If HandleError() = vbRetry Then Resume
End Sub

