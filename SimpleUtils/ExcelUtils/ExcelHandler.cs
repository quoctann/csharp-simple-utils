using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// Excel Object Library
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Diagnostics;
using System.IO;

namespace ExcelUtils
{   
    public class ExcelHandler : IDisposable
    {
        #region private variables
        private static Excel.Application App = null;
        private static Excel.Workbook Workbook = null;
        private static Excel.Worksheet Worksheet = null;
        private static Excel.Range Range = null;
        #endregion

        /// <summary>
        /// Open an excel file by filepath on local machine
        /// </summary>
        /// <param name="path">Excel file path</param>
        /// <returns>Workbook object that contains worksheets</returns>
        public static Excel.Workbook Open(string path)
        {
            try
            {
                if (!File.Exists(path))
                    throw new FileNotFoundException("Excel file path uncorrect");

                // Instance of Excel object
                App = new Excel.Application()
                {
                    Visible = false,
                };

                Workbook = App.Workbooks.Open(path);
            }
            catch (Exception ex)
            {
                App = null;
                Debug.Print(ex.Message);
            }

            return Workbook;
        }

        /// <summary>
        /// Release COM object
        /// </summary>
        /// <param name="obj">COM object need to release</param>
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

        /// <summary>
        /// Kill excel instance running on background
        /// </summary>
        public void Dispose()
        {
            ReleaseComObject(App);
            ReleaseComObject(Workbook);
            ReleaseComObject(Worksheet);
        }
    }
}
