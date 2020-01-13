using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using MyntraExcelAddin.Constant;

namespace MyntraExcelAddin.Service
{
    public class SheetUpdater
    {
        public Excel._Worksheet sheet;
        
        public SheetUpdater(Excel._Worksheet sheet)        
        {
            this.sheet = sheet;
        }
        
        public void HandoverIdsUpdate(List<long> handoverIds)
        {
            int row = 2;
            foreach(long id in handoverIds)
            {
                sheet.Cells[row, ColumnName.handoverId].Value = id;
                ++row;
            }
        }

    }
}
