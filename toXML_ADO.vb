Sub toXML()
    Application.ScreenUpdating = False

    Dim sBuff() As String
    Dim tmpStr As String
    Dim targetHeader As String
    Dim activeRowIndex As Integer
    Dim activeColIndex As Integer

    ' ヘッダーの合計業
    Dim headerTotalRows As Integer: headerTotalRows = 2
    ' ヘッダー名参照行の指定
    Dim headerRow As Integer: headerRow = 2
    ' 左端から無視する列数
    Dim ignoreCols As Integer: ignoreCols = 3
    ' ルート要素のタグ名
    Dim wrapper As String: wrapper = "DocumentElement"
    ' 各データを囲むタグ名
    Dim container As String: container = "MDHLP"
    ' 文字コード
    Dim charCode As String: charCode = "UTF-8" ' or Shift_JIS
    ' 保存先
    Dim savePath As String: savePath = Application.ThisWorkbook.Path & "/" & Application.ThisWorkbook.Name & ".xml"

    Dim lastRowIndex As Integer: lastRowIndex = Cells(Rows.Count, 1).End(xlUp).Row
    Dim lastColIndex As Integer: lastColIndex = Cells(1, Columns.Count).End(xlToLeft).Column

    ' header: table(headerRow, [列])
    Dim table(): table = Range(Cells(1, ignoreCols + 1), Cells(lastRowIndex, lastColIndex))

    ' XMLを塊ごとに格納するための配列
    ReDim sBuff(1 To UBound(table, 1) * UBound(table, 2) * 2)

    ' XML組み立て
    Dim xml As String
    xml = "<?xml version=""1.0"" encoding=""" & charCode & """ standalone=""yes""?>" & vbCrLf
    xml = xml &  "<" & wrapper & ">" & vbCrLf

    Dim i As Integer: i = 1
    For activeRowIndex = (headerTotalRows + 1) To lastRowIndex
        sBuff(i) = Space(2) & "<" & container & ">" & vbCrLf
        i = i + 1
        For activeColIndex = (ignoreCols + 1) To lastColIndex
          targetHeader = table(headerRow, activeColIndex - ignoreCols)
          tmpStr = tmpStr & Space(4) & "<" & targetHeader & ">"
          tmpStr = tmpStr & Cells(activeRowIndex, activeColIndex).Value
          tmpStr = tmpStr & Space(2) & "</" & targetHeader & ">" & vbCrLf
          sBuff(i) = tmpStr
          tmpStr = ""
          i = i + 1
        Next activeColIndex
        sBuff(i) = Space(2) & "</" & container & ">" & vbCrLf
        i = i + 1
    Next activeRowIndex
    xml = xml & Join(sBuff, "")
    xml = xml & "</" & wrapper & ">"

    ' ADOライブラリの読み込み
    Dim ado As Object: Set ado = CreateObject("ADODB.Stream")
    ado.Open
    ado.Type = 2 ' adTypeText
    ado.Charset = charCode
    ado.WriteText xml, 0 ' adWriteChar
    ado.SaveToFile savePath, 2 ' adSaveCreateOverWrite
    ado.Close
    Set ado = Nothing

End Sub
