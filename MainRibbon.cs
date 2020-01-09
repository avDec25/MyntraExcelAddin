using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using MyntraExcelAddin.Service;
using MyntraExcelAddin.SystemObjects;
using MyntraExcelAddin.Entity;

namespace MyntraExcelAddin
{
    public partial class MainRibbon
    {
        public Excel._Workbook xlWorkbook;
        public Excel._Worksheet syssheet;
        public Excel._Worksheet sheet;
        public ExternalServiceMessenger messenger;        

        public void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            messenger = new ExternalServiceMessenger();
        }

        private void GetTemplate_Click(object sender, RibbonControlEventArgs e)
        {            
            xlWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            // deletes all unnecessary sheets
            try
            {
                while (xlWorkbook.ActiveSheet != null)
                {
                    xlWorkbook.ActiveSheet.Delete();
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                syssheet = xlWorkbook.ActiveSheet;
                sheet = xlWorkbook.Worksheets.Add();
                syssheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            }

            SheetDecorator decorator = new SheetDecorator(messenger, sheet,syssheet);
            decorator.SetDropDowns();
            decorator.GenerateHeader();

            Validate.Enabled = true;
            UploadSheet.Enabled = true;
            GetTemplate.Enabled = false;
        }

        private void Validate_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = sheet.get_Range("A1", last);
            int lastUsedRow = last.Row;

            List<int> rows = new List<int>();
            for(int i = 2; i <= lastUsedRow; ++i)
            {
                rows.Add(i);
            }

            DataExtractor extractor = new DataExtractor(sheet);
            List<Handover> handoverlist = extractor.ExtractHandovers(rows);

            NotificationService notify = new NotificationService();

            //if (handoverlist == null) { 
            //    return; 
            //}

            notify.ValidationComplete();

            //DataValidator validator = new DataValidator(sheet);
            //var handoverReport = validator.ValidateHandovers(handoverlist);
        }
    }
}
