
Sub toXML()

	Application.ScreenUpdating = False

	Dim sname As String
	Dim headerRows As Integer
	Dim lastRowIndex As Integer
	Dim lastColIndex As Integer
	Dim ignoreCols As Integer
	Dim activeRowIndex As Integer
	Dim activeColIndex As Integer

	headerRows = 2
	lastRowIndex = Cells(Rows.Count, 1).End(xlUp).Row
	lastColIndex = Cells(1, Columns.Count).End(xlToLeft).Column
	ignoreCols = 3

	For activeRowIndex = (headerRows + 1) To lastRowIndex
		For activeColIndex = (ignoreCols + 1) To lastColIndex
			Cells(activeRowIndex, activeColIndex).Value = 1
		Next activeColIndex
	Next activeRowIndex

End Sub


