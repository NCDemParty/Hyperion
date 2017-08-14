Sub CreateSlides()
    '*** Original Sourcecode taken from http://superuser.com/questions/323408/excel-data-into-powerpoint-slides ***
    'Open the Excel workbook. Change the filename here.
    Dim OWB As New Excel.Workbook
    Set OWB = Excel.Application.Workbooks.Open("C:\Users\JessePresnell\Desktop\test.xlsx")
    'Grab the first Worksheet in the Workbook
    Dim WS As Excel.Worksheet
    Dim sCurrentText As String
    Dim oSl As Slide
    Dim oSh As Shape
    Set WS = OWB.Worksheets(1)
    Dim i As Long
    'Loop through each used row in Column A
    For i = 1 To WS.Range("A65536").End(xlUp).Row
        'Copy the first slide and paste at the end of the presentation
        ActivePresentation.Slides(1).Copy
        Set oSl = ActivePresentation.Slides(1).Duplicate.Item(1)
        sCurrentText = WS.Cells(i, 1).Value
        ' find each shape with "@COL1@" in text, replace it with value from worksheet
        For Each oSh In oSl.Shapes
          ' Make sure the shape can hold text and if is, that it IS holding text
          If oSh.HasTextFrame Then
            If oSh.TextFrame.HasText Then
              ' it's got text, do the replace
              With oSh.TextFrame.TextRange
                .Replace "@COL1@", sCurrentText
              End With
            End If
          End If
        Next
    Next
End Sub
