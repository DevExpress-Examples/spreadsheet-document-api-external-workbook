using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocServerExternalWorkbookSample
{
    public partial class Form1 : Form
    {
        Workbook myWorkbook = new Workbook();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            #region #addexternalworkbook
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
            #endregion #addexternalworkbook
            button1.Enabled = !button1.Enabled;
        }


        DataTable CreateDataTable(int rowCount)
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

        private void button2_Click(object sender, EventArgs e)
        {
            #region #insertexternalreference
            if (myWorkbook.ExternalWorkbooks.Count == 0)
            {
                return;
            }
            IWorkbook extWorkbook = (IWorkbook)myWorkbook.ExternalWorkbooks[0];
            string extWorkbookName = extWorkbook.Options.Save.CurrentFileName;
            string sFormula = String.Format("=[{0}]Sheet1!A1", extWorkbookName);
            myWorkbook.Worksheets[0].Cells["A1"].Formula = sFormula;
            myWorkbook.SaveDocument("Test.xlsx");
            System.Diagnostics.Process.Start("Test.xlsx");
            #endregion #insertexternalreference
        }
    }
}
