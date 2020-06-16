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
    Dim wrapper As String
    Dim container As String
    Dim charCode As String
    Dim savePath As String
    Dim i As Integer

    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ' ヘッダーの合計業
    headerTotalRows = 2
    ' ヘッダー名参照行の指定
    headerRow = 2
    ' 左端から無視する列数
    ignoreCols = 3
    ' ルート要素のタグ名
    wrapper = "Wrapper"
    ' 各データを囲むタグ名
    container = "Container"
    ' 文字コード
    charCode = "UTF-8" ' Shift_JIS
    ' 保存先
    savePath = Application.ThisWorkbook.Path & "/" & Application.ThisWorkbook.Name & ".xml"


    lastRowIndex = Cells(Rows.Count, 1).End(xlUp).Row
    lastColIndex = Cells(1, Columns.Count).End(xlToLeft).Column

    table = Range(Cells(1, ignoreCols + 1), Cells(lastRowIndex, lastColIndex))
    ' header: table(headerRow, [?h??])

    ReDim sBuff(1 To UBound(table, 1) * UBound(table, 2) * 2)

    xml = "<?xml version=""1.0"" encoding=""" & charCode & """?>" & vbCrLf
    xml = xml & "<" & wrapper & ">" & vbCrLf
    i = 1
    For activeRowIndex = (headerTotalRows + 1) To lastRowIndex
        sBuff(i) = "<" & container & ">" & vbCrLf & vbTab & "<RowNum>" & activeRowIndex - headerTotalRows & "</RowNum>" & vbCrLf
        i = i + 1
        For activeColIndex = (ignoreCols + 1) To lastColIndex
          targetHeader = table(headerRow, activeColIndex - ignoreCols)
          tmpStr = tmpStr & vbTab & "<" & targetHeader & ">"
          tmpStr = tmpStr & Cells(activeRowIndex, activeColIndex).Value
          tmpStr = tmpStr & "</" & targetHeader & ">" & vbCrLf
          sBuff(i) = tmpStr
          tmpStr = ""
          i = i + 1
        Next activeColIndex
        sBuff(i) = "</" & container & ">" & vbCrLf
        i = i + 1
    Next activeRowIndex
    xml = xml & Join(sBuff, "")
    xml = xml & "</" & wrapper & ">"

    ado.Open
    ado.Type = 2 ' adTypeText
    ado.Charset = charCode
    ado.WriteText xml, 0 ' adWriteChar
    ado.SaveToFile savePath, 2 ' adSaveCreateOverWrite
    ado.Close

    Set ado = Nothing

End Sub
