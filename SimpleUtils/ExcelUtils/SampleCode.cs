using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelUtils
{
    public class SampleCode
    {
        /// <summary>
        /// The code accepts an Excel file to open, iterates thru the
        /// WorkSheets collection checking the name passed in parameter 2
        /// to see if it matches and if so sets a boolean so after the for-next
        /// we can know if the sheet is there and safe to work on.
        /// </summary>
        /// <param name="FileName"></param>
        /// <param name="SheetName"></param>
        public void OpenExcel(string FileName, string SheetName)
        {
            List<string> SheetNames = new List<string>();

            bool Proceed = false;
            Excel.Application xlApp = null;
            Microsoft.Office.Interop.Excel.Workbooks xlWorkBooks = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Sheets xlWorkSheets = null;

            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            xlWorkBooks = xlApp.Workbooks;
            xlWorkBook = xlWorkBooks.Open(FileName);

            xlApp.Visible = false;
            xlWorkSheets = xlWorkBook.Sheets;

            for (int x = 1; x <= xlWorkSheets.Count; x++)
            {
                xlWorkSheet = (Excel.Worksheet)xlWorkSheets[x];

                SheetNames.Add(xlWorkSheet.Name);

                if (xlWorkSheet.Name == SheetName)
                {
                    Proceed = true;
                    Excel.Range xlRange1 = null;
                    xlRange1 = xlWorkSheet.Range["A1"];
                    xlRange1.Value = "Hello";
                    Marshal.FinalReleaseComObject(xlRange1);
                    xlRange1 = null;
                    xlWorkSheet.SaveAs(FileName);
                    break;
                }

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkSheet);
                xlWorkSheet = null;
            }

            xlWorkBook.Close();
            xlApp.Quit();

            ReleaseComObject(xlWorkSheets);
            ReleaseComObject(xlWorkSheet);
            ReleaseComObject(xlWorkBook);
            ReleaseComObject(xlWorkBooks);
            ReleaseComObject(xlApp);

            if (Proceed)
            {
                MessageBox.Show("Found sheet, do your work here.");
            }
            else
            {
                MessageBox.Show("Sheet not located");
            }

            MessageBox.Show("Sheets available \n" + String.Join("\n", SheetNames.ToArray()));
        }

        private void ReleaseComObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
        }
    }
}
