Attribute VB_Name = "SSOVCoreModule"
'This module contains this program's core.
Option Explicit

'This procedure displays the objects/events/parameters in the selected script/objecttype.
Public Sub DisplayObjects(GetEventsParameters As Boolean, EventParameterList As ListBox, ObjectsList As ListBox)
On Error GoTo ErrorTrap
Dim Character As String
Dim Code As String
Dim CodeLine As String
Dim CurrentObject As String
Dim InComment As Boolean
Dim InFileName As Boolean
Dim InString As Boolean
Dim Offset As Long
Dim Position As Long
Dim PreviousCharacter As String
Dim WasInComment As Boolean

   Code = ScriptCode()
   EventParameterList.Clear
   InComment = False
   InFileName = False
   InString = False

   If Not GetEventsParameters Then ObjectsList.Clear

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
            CodeLine = Mid$(Code, Offset, (Position - Offset) + 1)
            CodeLine = Trim$(Replace(CodeLine, vbCr, " "))
        
            If Left$(CodeLine, Len("objectType ")) = "objectType " Or Left$(CodeLine, Len("sub ")) = "sub " Then
               CurrentObject = Trim$(Left$(CodeLine, InStr(CodeLine, "(") - 1))
               If Not GetEventsParameters Then ObjectsList.AddItem CurrentObject
            End If
      
            If GetEventsParameters Then
               If Left$(CurrentObject, Len("objectType ")) = "objectType " And CurrentObject = ObjectsList.List(ObjectsList.ListIndex) Then
                  If Left$(CodeLine, Len("event ")) = "event " Then
                     If Right$(CodeLine, 1) = "{" Then CodeLine = Trim$(Left$(CodeLine, Len(CodeLine) - 1))
                     EventParameterList.AddItem CodeLine
                  End If
               End If
          
               If Left$(CurrentObject, Len("sub ")) = "sub " And CurrentObject = ObjectsList.List(ObjectsList.ListIndex) Then
                  If Left$(CodeLine, Len("sub ")) = "sub " Then
                     CodeLine = Trim$(Mid$(CodeLine, InStr(CodeLine, "(") + 1))
                     If InStr(CodeLine, ")") > 0 Then CodeLine = Left$(CodeLine, InStr(CodeLine, ")") - 1)
                        EventParameterList.AddItem CodeLine
                     End If
                  End If
               End If
               Offset = 0
            End If
         End If
      End If
   Next Position
Endroutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume Endroutine
End Sub

'This procedure handles any errors that occur.
Public Sub HandleError()
Dim ErrorCode As Long
Dim Message As String

   ErrorCode = Err.Number
   Message = Err.Description

   On Error Resume Next
   MsgBox "Error: " & ErrorCode & vbCr & Message, vbExclamation
End Sub


'This procedure quits this program after user confirmation.
Public Function Quit() As Long
On Error Resume Next
Dim Choice As Long

   Choice = MsgBox("Do you want to quit?", vbYesNo Or vbQuestion Or vbDefaultButton2)
   If Choice = vbYes Then Unload SSOVBox

   Quit = Choice
End Function



'This procedure manages the current script file code.
Public Function ScriptCode(Optional ScriptFile As String = Empty) As String
On Error GoTo ErrorTrap
Dim FileHandle As Integer
Static Code As String

   If Not ScriptFile = Empty Then
      If Not Dir$(ScriptFile, vbNormal) = Empty Then
         FileHandle = FreeFile()
         Open ScriptFile For Binary Lock Read Write As FileHandle
            Code = Trim$(Input$(LOF(FileHandle), FileHandle))
         Close FileHandle
         
         Code = Replace(Code, vbLf, Empty)
         Code = Replace(Code, vbTab, Empty)
         Code = Replace(Code, "  ", " ")
      End If
   End If

Endroutine:
   ScriptCode = Code
   Exit Function

ErrorTrap:
   HandleError
   Resume Endroutine
End Function


'This procedure sorts the items in the specified list by name.
Public Sub SortByName(List As Object)
On Error GoTo ErrorTrap
Dim Index As Long
Dim OtherIndex As Long

   With List
      For Index = 0 To .ListCount - 1
         For OtherIndex = 0 To .ListCount - 1
            If .List(Index) < .List(OtherIndex) Then SwapItems List, Index, OtherIndex
         Next OtherIndex
      Next Index
   End With
Endroutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume Endroutine
End Sub


'This procedure sorts the items in the specified list by type.
Public Sub SortByType(List As Object)
On Error GoTo ErrorTrap
Dim Index As Long
Dim OtherIndex As Long

   With List
      For Index = 0 To .ListCount - 1
         For OtherIndex = 0 To .ListCount - 1
            If Left$(.List(Index), InStr(.List(Index), " ")) < Left$(.List(OtherIndex), InStr(.List(OtherIndex), " ")) Then SwapItems List, Index, OtherIndex
         Next OtherIndex
      Next Index
   End With
Endroutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume Endroutine
End Sub


'This procedure exchanges the two specified list items with each other.
Private Sub SwapItems(List As Object, Index1 As Variant, Index2 As Variant)
On Error GoTo ErrorTrap
   Dim Temporary As Variant
   
   With List
      Temporary = .List(Index1)
      .List(Index1) = .List(Index2)
      .List(Index2) = Temporary
   End With
Endroutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume Endroutine
End Sub


