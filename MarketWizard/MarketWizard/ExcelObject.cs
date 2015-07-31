using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace MarketWizard
{
    class ExcelObject
    {
        private ExcelApp excelApp;
        private Excel._Worksheet excelWorksheet;
        private Excel.Workbook excelWorkbook;
        private Dictionary<int, String> addressMap;
        private int sentCol = 0;

        public ExcelObject(String sheetPath)
        {
            constructExcelObject(sheetPath);
        }

        private void constructExcelObject(String sheetPath) 
        {
            excelApp = new Excel.Application();
            excelWorkbook = excelApp.Workbooks.Open(@sheetPath);
            excelWorksheet = excelWorkbook.Sheets[1];
            Excel.Range excelRange = excelWorksheet.UsedRange;

            addressMap = new Dictionary<int, string>();

            fillAddressList(excelWorkbook, excelWorksheet, excelRange);
        }

        private void fillAddressList(Excel.Workbook excelWorkbook, Excel._Worksheet excelWorksheet, Excel.Range excelRange)
        {
            
            int rowCount = excelRange.Rows.Count;
            int colCount = excelRange.Columns.Count;

            int emailColumn = 0;

            // Get the column that holds email addresses
            for (int potCol = 1; potCol <= colCount; potCol++)
            {
                String currentColHeading = excelRange.Cells[1, potCol].Value2.ToString();
                if (currentColHeading.Equals("Email"))
                {
                    emailColumn = potCol;
                }

                // Get the Sent column
                if (currentColHeading.Equals("Sent"))
                {
                    sentCol = potCol;
                }
            }

            if (sentCol == 0)
            {
                MessageBox.Show("The specified address worksheet cannot be written to.", "Warning!");
            }

            // Now, read all of the data from that column into the list
            for (int newRow = 2; newRow <= rowCount; newRow++)
            {
                String newAddr = excelRange.Cells[newRow, emailColumn].Value2.ToString();
                addressMap.Add(newRow, newAddr);
            }
        }

        public Dictionary<int, String> getAddresses()
        {
            return addressMap;
        }

        public void logSuccess(int row)
        {
            if (sentCol != 0)
            {
                excelWorksheet.Cells[row, sentCol] = DateTime.Today;
            }
            
        }

        public void saveAndClose()
        {
            excelWorkbook.Save();
            excelWorkbook.Close();
        }
         
    }
}
