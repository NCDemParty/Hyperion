Sub Button_Click()
    ActiveWorkbook.Sheets.Add Before:=Worksheets(Worksheets.Count)
    i = Sheet1.Cells(8, 6).Value
    CandidateCount = Sheet1.Cells(11, 8).Value
    CandidateStart = CandidateCount + 1
    SheetName = Sheet1.Cells(13, 8).Value
    OrSheetName = Sheet1.Cells(13, 8).Value
    v = 0
    For Each Sheet In Worksheets
        If SheetName = Sheet.Name Then
            SheetName = OrSheetName & "_Ballot" & v
            v = v + 1
        End If
    Next Sheet
    ActiveSheet.Name = SheetName
    
    ActiveSheet.Cells(CandidateCount + 3, 1).Value = "Delegates Checked/Voted"
    For x = 1 To i
    b = 15 + x
        Subdistrict = Sheet1.Cells(b, 2).Value
        MembersPresent = Sheet1.Cells(b, 6).Value
        ActiveSheet.Cells(CandidateCount + 3, x + 1).Formula = _
        "=(" & MembersPresent & " - " & Replace(Cells(1, x + 1).Address(0, 0), 1, "") & CandidateCount + 2 & ")"
        ActiveSheet.Cells(1, (x + 1)).Value = Subdistrict
        For a = 2 To CandidateCount + 1
            ActiveSheet.Cells(a, x + 1).Value = 0
        Next a
    Next x
    For b = 1 To CandidateCount
        Named = "Candidate Name" & " " & b
        ActiveSheet.Cells((b + 1), 1).Value = Named
    Next b
    ActiveSheet.Cells(CandidateCount + 2, 1).Value = "Raw Vote Totals"
    For m = 2 To i + 2
        ActiveSheet.Cells(CandidateCount + 2, m).Formula = _
        "=SUM(" & ActiveSheet.Range(ActiveSheet.Cells(2, m), ActiveSheet.Cells(CandidateStart, m)).Address(False, False) & ")"
    Next m
    For a = 2 To CandidateCount + 1
        ActiveSheet.Cells(a, i + 2).Formula = _
        "=SUM(" & ActiveSheet.Range(ActiveSheet.Cells(2, 2), ActiveSheet.Cells(2, i + 1)).Address(False, False) & ")"
    Next a
    StartWeight = 1 + (CandidateCount + 5)
    For x = 1 To i
        b = 15 + x
        VotesPerRep = Sheet1.Cells(b, 5).Value
        ColLtr = Replace(Cells(1, x + 1).Address(0, 0), 1, "")
        ActiveSheet.Cells(StartWeight + 1, (x + 1)).Value = ActiveSheet.Cells(1, (x + 1)).Value
        For co = 2 To CandidateCount + 1
            rawvote = ActiveSheet.Cells(co, x + 1).Value
            ActiveSheet.Cells(co + StartWeight, x + 1).BorderAround xlContinuous, xlThin
            ActiveSheet.Cells(co + StartWeight, x + 1).Value = "=(" & VotesPerRep & "*" & ColLtr & co & ")"
        Next co
    Next x
    For b = StartWeight + 1 To CandidateCount + StartWeight + 1
        ActiveSheet.Cells((b + 1), 1).Value = "=(A" & b - (CandidateCount + 5) & ")"
    Next b
    ActiveSheet.Cells(CandidateCount + StartWeight + 2, 1).Value = "Weighted Vote Totals"
    For m = 2 To i + 2
        ActiveSheet.Cells(CandidateCount + 2 + StartWeight, m).Formula = _
        "=SUM(" & ActiveSheet.Range(ActiveSheet.Cells(StartWeight + 2, m), ActiveSheet.Cells(CandidateStart + StartWeight, m)).Address(False, False) & ")"
    Next m
    For a = StartWeight + 2 To CandidateCount + StartWeight + 1
        ActiveSheet.Cells(a, i + 2).BorderAround xlContinuous, xlThin
        ActiveSheet.Cells(a, i + 2).Formula = _
        "=SUM(" & ActiveSheet.Range(ActiveSheet.Cells(StartWeight + 2, 2), ActiveSheet.Cells(StartWeight + 2, i + 1)).Address(False, False) & ")"
    Next a
    For x = 1 To i + 2
        For b = 1 To CandidateCount + 3:
            ActiveSheet.Cells(b, x).BorderAround xlContinuous, xlThin
            ActiveSheet.Cells(StartWeight + b, x).BorderAround xlContinuous, xlThin
        Next b
    Next x
    With ActiveSheet.Columns("A")
        .ColumnWidth = .ColumnWidth * 3
    End With
    ActiveSheet.Range("B1:Z1").Columns.AutoFit
End Sub
