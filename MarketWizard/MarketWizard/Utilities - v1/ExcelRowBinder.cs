using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarketWizard.Utilities
{
    // Small utility class that binds a row to a particular value in the row
    class ExcelRowBinder
    {
        private int rowNumber;
        private String cellContents;

        public ExcelRowBinder(int row, String value)
        {
            rowNumber = row;
            cellContents = value;
        }

        public int getRowNumber()
        {
            return rowNumber;
        }

        public String getValue()
        {
            return cellContents;
        }
    }
}
