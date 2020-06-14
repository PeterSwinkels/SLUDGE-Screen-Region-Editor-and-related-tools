Attribute VB_Name = "SSCModule"
'This module contains the SLUDGE Scale Calculator.
Option Explicit

'This procedure calculates dividers based on the user's input.
Public Sub Main()
On Error GoTo ErrorTrap
Dim Choice As Long
Dim Divider As Long
Dim GenerateSetScaleStatement As Boolean
Dim Horizon As Long
Dim ObjectScale As Double
Dim PutOnClipBoard As Boolean
Dim YPosition As Long

   Divider = 0
   GenerateSetScaleStatement = False
   Horizon = 0
   ObjectScale = 0
   PutOnClipBoard = False
   YPosition = 0
   
   Choice = MsgBox("Place the resulting dividers on clipboard?", vbYesNo Or vbQuestion)
   If Choice = vbYes Then
      PutOnClipBoard = True
      Choice = MsgBox("Generate a setScale statement containing the divider?", vbYesNo Or vbQuestion)
      If Choice = vbYes Then GenerateSetScaleStatement = True
   End If
     
   Do
      Horizon = Val(InputBox$("Specify the horizon's position.", , Horizon))
      YPosition = Val(InputBox$("Specify the vertical position for an object.", , YPosition))
      ObjectScale = Val(InputBox$("How many times the actual size should the object appear?", , ObjectScale))
      If Horizon + YPosition + ObjectScale = 0 Then End
      Divider = (YPosition - Horizon) / ObjectScale
   
      If PutOnClipBoard Then
         Clipboard.Clear
         If GenerateSetScaleStatement Then
            Clipboard.SetText "setScale (" & CStr(Horizon) & ", " & CStr(Divider) & ");", vbCFText
         ElseIf Not GenerateSetScaleStatement Then
            Clipboard.SetText CStr(Divider), vbCFText
         End If
      End If
   
      Choice = MsgBox("(Y - Horizon) / Scale = " & CStr(Divider) & " (divider)" & vbCr & vbCr & "Calculate new divider?", vbYesNo Or vbInformation)
   Loop Until Choice = vbNo
EndRoutine:
   Exit Sub

ErrorTrap:
   MsgBox "Error: " & Err.Number & vbCr & Err.Description, vbExclamation
   Resume EndRoutine
End Sub

