using DevExpress.Spreadsheet;
using System.Data;
using System.Diagnostics;

namespace SpreadsheetApiExternalWorkbook
{
    internal class Program
    {

        static void Main(string[] args)
        {
            Workbook myWorkbook = new Workbook();

            Workbook externalWorkbook = new Workbook();
            externalWorkbook.Options.Save.CurrentFileName = "ExternalDocument.xlsx";
            // Check whether the external workbook is already referenced.
            foreach (IWorkbook item in myWorkbook.ExternalWorkbooks)
            {
                if (item.Options.Save.CurrentFileName == externalWorkbook.Options.Save.CurrentFileName)
                    return;
            }
            externalWorkbook.Worksheets[0].Import(CreateDataTable(10), false, 0, 0);
            externalWorkbook.SaveDocument("ExternalDocument.xlsx");
            myWorkbook.ExternalWorkbooks.Add(externalWorkbook);

            if (myWorkbook.ExternalWorkbooks.Count == 0)
            {
                return;
            }
            IWorkbook extWorkbook = (IWorkbook)myWorkbook.ExternalWorkbooks[0];
            string extWorkbookName = extWorkbook.Options.Save.CurrentFileName;
            string sFormula = String.Format("=[{0}]Sheet1!A1", extWorkbookName);
            myWorkbook.Worksheets[0].Cells["A1"].Formula = sFormula;
            myWorkbook.SaveDocument("Test.xlsx");
            Process.Start(new ProcessStartInfo("Test.xlsx") { UseShellExecute = true });
        }
        static DataTable CreateDataTable(int rowCount)
        {
            DataTable someDT = new DataTable();
            for (int i = 0; i < 5; i++)
            {
                someDT.Columns.Add("Value" + i.ToString(), typeof(int));
            }
            Random myRand = new Random();
            for (int i = 0; i < rowCount; i++)
            {
                someDT.Rows.Add(myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100));
            }
            return someDT;
        }

    }
}
