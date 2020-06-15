Sub toXML()

    Application.ScreenUpdating = False

    Dim sname As String
    Dim headerTotalRows As Integer
    Dim lastRowIndex As Integer
    Dim lastColIndex As Integer
    Dim ignoreCols As Integer
    Dim activeRowIndex As Integer
    Dim activeColIndex As Integer
    Dim oneRowCharCount As Integer
    Dim headerRow As Integer
    Dim table()
    Dim sBuff() As String
    Dim tmpStr As String
    Dim targetHeader As String
    Dim xml As String
    Dim i As Integer

    headerTotalRows = 2
    headerRow = 2
    lastRowIndex = Cells(Rows.Count, 1).End(xlUp).Row
    lastColIndex = Cells(1, Columns.Count).End(xlToLeft).Column
    ignoreCols = 3

    table = Range(Cells(1, ignoreCols + 1), Cells(lastRowIndex, lastColIndex))
    ' header: table(headerRow, [数字])

    ReDim sBuff(1 To UBound(table, 1) * UBound(table, 2) * 2)

    xml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    xml = xml & "<ExcelField>" & vbCrLf
    i = 1
    For activeRowIndex = (headerTotalRows + 1) To lastRowIndex
        sBuff(i) = "<Data>" & vbCrLf & "<RowNum>" & activeRowIndex - headerTotalRows & "</RowNum>" & vbCrLf
        i = i + 1
        For activeColIndex = (ignoreCols + 1) To lastColIndex
          targetHeader = table(headerRow, activeColIndex - ignoreCols)
          tmpStr = tmpStr & "<" & targetHeader & ">"
          tmpStr = tmpStr & Cells(activeRowIndex, activeColIndex).Value
          tmpStr = tmpStr & "</" & targetHeader & ">" & vbCrLf
          sBuff(i) = tmpStr
          tmpStr = ""
          i = i + 1
        Next activeColIndex
        sBuff(i) = "</Data>" & vbCrLf
        i = i + 1
    Next activeRowIndex
    xml = xml & Join(sBuff, "")


    xml = xml & "</ExcelField>"
    Cells(1, 1).Value = xml

End Sub


