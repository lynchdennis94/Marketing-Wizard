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
        private Dictionary<int, Utilities.ExcelRowBinder> addressMap;
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

            addressMap = new Dictionary<int, Utilities.ExcelRowBinder>();

            fillAddressList(excelWorkbook, excelWorksheet, excelRange);
        }

        private void fillAddressList(Excel.Workbook excelWorkbook, Excel._Worksheet excelWorksheet, Excel.Range excelRange)
        {
            bool haveEmail=false, haveSent=false;
            int rowCount = excelRange.Rows.Count;
            int colCount = excelRange.Columns.Count;

            int emailColumn = 0;

            // Get the column that holds email addresses
            for (int potCol = 1; potCol <= colCount; potCol++)
            {
                if (haveEmail && haveSent)
                {
                    // We already have all of the information we need, so we don't need to worry
                    // about scanning potentially harmful columns
                    break;
                }
                try
                {
                    String currentColHeading = excelRange.Cells[1, potCol].Value2.ToString();

                    if (currentColHeading.ToLower().Trim().Equals("email"))
                    {
                        emailColumn = potCol;
                        haveEmail = true;
                    }

                    // Get the Sent column
                    if (currentColHeading.ToLower().Trim().Equals("sent"))
                    {
                        sentCol = potCol;
                        haveSent = true;
                    }
                }
                catch (Exception e)
                {
                    if (haveEmail && !haveSent)
                    {
                        // Since we don't have both but have email, we can continue
                        break;
                    }
                    else if (haveSent && !haveEmail)
                    {
                        // We need to send a message box and return to the window
                        MessageBox.Show("There was an error importing addresses."
                                    + "\nPlease ensure that there is an 'Email' column, and try again.",
                                    "Address Importation Error");
                        return;
                    }
                    else
                    {
                        // We need to send a message box and return to the window
                        MessageBox.Show("There was an error importing addresses."
                                    + "\nPlease ensure that there is an 'Email' and 'Sent' column, and try again.",
                                    "Address Importation Error");
                        return;
                    }
                }  
            }

            if (sentCol == 0)
            {
                MessageBox.Show("The specified address worksheet cannot be written to.", "Warning!");
            }

            // Now, read all of the data from that column into the list
            int indexKey = 0;
            for (int newRow = 2; newRow <= rowCount; newRow++)
            {
                String newAddrStr = excelRange.Cells[newRow, emailColumn].Value2.ToString();
                String[] newAddrArray = newAddrStr.Split(','); // Account for multiple emails per cell
                foreach (String newAddr in newAddrArray)
                {
                    if (newAddr != "")
                    {
                        addressMap.Add(indexKey, new Utilities.ExcelRowBinder(newRow, newAddr));
                        indexKey++;
                    }
                }
                
            }
        }

        public Dictionary<int, Utilities.ExcelRowBinder> getAddresses()
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
