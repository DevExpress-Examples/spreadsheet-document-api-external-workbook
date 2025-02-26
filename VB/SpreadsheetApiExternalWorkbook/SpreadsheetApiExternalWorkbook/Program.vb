Imports DevExpress.Spreadsheet
Imports System.Data
Imports System.Diagnostics

Namespace SpreadsheetApiExternalWorkbook
	Friend Class Program

		Shared Sub Main(ByVal args() As String)
			Dim myWorkbook As New Workbook()

			Dim externalWorkbook As New Workbook()
			externalWorkbook.Options.Save.CurrentFileName = "ExternalDocument.xlsx"
			' Check whether the external workbook is already referenced.
			For Each item As IWorkbook In myWorkbook.ExternalWorkbooks
				If item.Options.Save.CurrentFileName = externalWorkbook.Options.Save.CurrentFileName Then
					Return
				End If
			Next item
			externalWorkbook.Worksheets(0).Import(CreateDataTable(10), False, 0, 0)
			externalWorkbook.SaveDocument("ExternalDocument.xlsx")
			myWorkbook.ExternalWorkbooks.Add(externalWorkbook)

			If myWorkbook.ExternalWorkbooks.Count = 0 Then
				Return
			End If
			Dim extWorkbook As IWorkbook = DirectCast(myWorkbook.ExternalWorkbooks(0), IWorkbook)
			Dim extWorkbookName As String = extWorkbook.Options.Save.CurrentFileName
			Dim sFormula As String = String.Format("=[{0}]Sheet1!A1", extWorkbookName)
			myWorkbook.Worksheets(0).Cells("A1").Formula = sFormula
			myWorkbook.SaveDocument("Test.xlsx")
			Process.Start(New ProcessStartInfo("Test.xlsx") With {.UseShellExecute = True})
		End Sub
		Private Shared Function CreateDataTable(ByVal rowCount As Integer) As DataTable
			Dim someDT As New DataTable()
			For i As Integer = 0 To 4
				someDT.Columns.Add("Value" & i.ToString(), GetType(Integer))
			Next i
			Dim myRand As New Random()
			For i As Integer = 0 To rowCount - 1
				someDT.Rows.Add(myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100))
			Next i
			Return someDT
		End Function

	End Class
End Namespace