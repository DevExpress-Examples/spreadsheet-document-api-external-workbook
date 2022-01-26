Imports DevExpress.Spreadsheet
Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms

Namespace DocServerExternalWorkbookSample

    Public Partial Class Form1
        Inherits Form

        Private myWorkbook As Workbook = New Workbook()

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
#Region "#addexternalworkbook"
            Dim externalWorkbook As Workbook = New Workbook()
            externalWorkbook.Options.Save.CurrentFileName = "ExternalDocument.xlsx"
            ' Check whether the external workbook is already referenced.
            For Each item As IWorkbook In myWorkbook.ExternalWorkbooks
                If Equals(item.Options.Save.CurrentFileName, externalWorkbook.Options.Save.CurrentFileName) Then Return
            Next

            externalWorkbook.Worksheets(0).Import(CreateDataTable(10), False, 0, 0)
            externalWorkbook.SaveDocument("ExternalDocument.xlsx")
            myWorkbook.ExternalWorkbooks.Add(externalWorkbook)
#End Region  ' #addexternalworkbook
            button1.Enabled = Not button1.Enabled
        End Sub

        Private Function CreateDataTable(ByVal rowCount As Integer) As DataTable
            Dim someDT As DataTable = New DataTable()
            For i As Integer = 0 To 5 - 1
                someDT.Columns.Add("Value" & i.ToString(), GetType(Integer))
            Next

            Dim myRand As Random = New Random()
            For i As Integer = 0 To rowCount - 1
                someDT.Rows.Add(myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100))
            Next

            Return someDT
        End Function

        Private Sub button2_Click(ByVal sender As Object, ByVal e As EventArgs)
#Region "#insertexternalreference"
            If myWorkbook.ExternalWorkbooks.Count = 0 Then
                Return
            End If

            Dim extWorkbook As IWorkbook = CType(myWorkbook.ExternalWorkbooks(0), IWorkbook)
            Dim extWorkbookName As String = extWorkbook.Options.Save.CurrentFileName
            Dim sFormula As String = String.Format("=[{0}]Sheet1!A1", extWorkbookName)
            myWorkbook.Worksheets(0).Cells("A1").Formula = sFormula
            myWorkbook.SaveDocument("Test.xlsx")
            System.Diagnostics.Process.Start("Test.xlsx")
#End Region  ' #insertexternalreference
        End Sub
    End Class
End Namespace
