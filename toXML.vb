Sub toXML()

    Application.ScreenUpdating = False

    Dim sname As String
    Dim headerRows As Integer
    Dim lastRowIndex As Integer
    Dim lastColIndex As Integer
    Dim ignoreCols As Integer
    Dim activeRowIndex As Integer
    Dim activeColIndex As Integer
    Dim oneRowCharCount As Integer
    Dim sBuff() As String
        Dim ary()
        Dim cellTo As Object
        Dim cellFrom As Object
    Dim i As Integer

    headerRows = 2
    lastRowIndex = Cells(Rows.Count, 1).End(xlUp).Row
    lastColIndex = Cells(1, Columns.Count).End(xlToLeft).Column
    ignoreCols = 3

        ' cellTo = Cells(1, ignoreCols)
        ' cellFrom = Cells(lastRowIndex, lastColIndex)

        ary = Range(Cells(1, ignoreCols), Cells(lastRowIndex, lastColIndex))

    oneRowCharCount = 0
    For i = ignoreCols To lastColIndex
        oneRowCharCount = oneRowCharCount + Len(Cells(1, i))
    Next i

        ReDim sBuff(1 To (oneRowCharCount * 1.2) * lastRowIndex)
        MsgBox (oneRowCharCount * 1.2) * lastRowIndex

    For activeRowIndex = (headerRows + 1) To lastRowIndex
        For activeColIndex = (ignoreCols + 1) To lastColIndex
            Cells(activeRowIndex, activeColIndex).Value = activeRowIndex * activeColIndex
        Next activeColIndex
    Next activeRowIndex

End Sub


