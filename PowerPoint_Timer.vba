Sub Time_Me()
Dim oshp As Shape
Dim oshpRng As ShapeRange
Dim osld As Slide
Dim oeff As Effect
Dim i As Integer
Dim Iduration As Integer
Dim Istep As Integer
Dim dText As Date
Dim texttoshow As String
On Error GoTo errhandler
If ActiveWindow.Selection.ShapeRange.Count > 1 Then
MsgBox "Please just select ONE shape!"
Exit Sub
End If
Set osld = ActiveWindow.Selection.SlideRange(1)
Set oshp = ActiveWindow.Selection.ShapeRange(1)
oshp.Copy

'change to suit
Istep = 1
Iduration = 120 'in seconds

For i = Iduration To 0 Step -Istep
Set oshpRng = osld.Shapes.Paste
With oshpRng
.Left = oshp.Left
.Top = oshp.Top
End With
dText = CDate(i \ 3600 & ":" & ((i Mod 3600) \ 60) & ":" & (i Mod 60))
If Iduration < 3600 Then
texttoshow = Format(dText, "Nn:Ss")
Else
texttoshow = Format(dText, "Hh:Nn:Ss")
End If
oshpRng(1).TextFrame.TextRange = texttoshow
Set oeff = osld.TimeLine.MainSequence _
.AddEffect(oshpRng(1), msoAnimEffectFlashOnce, , msoAnimTriggerAfterPrevious)
oeff.Timing.Duration = Istep
Next i
oshp.Delete
Exit Sub
errhandler:
MsgBox "**ERROR** - Maybe nothing is selected?"
End Sub
