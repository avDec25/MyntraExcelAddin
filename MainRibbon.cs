using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using MyntraExcelAddin.Service;
using MyntraExcelAddin.SystemObjects;
using MyntraExcelAddin.Entity;
using System.Windows.Forms;

namespace MyntraExcelAddin
{
    public partial class MainRibbon : IDisposable
    {
        public Excel._Workbook xlWorkbook;
        public Excel._Worksheet syssheet;
        public Excel._Worksheet sheet;
        public SheetDecorator decorator;
        public ExternalServiceMessenger messenger;
        public Excel.Application app;
        DataExtractor extractor;
        DataValidator validator;
        EventManagement eventmanager;
        ValueDeterminer determiner;

        public void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            messenger = new ExternalServiceMessenger();            
            //app = Globals.ThisAddIn.Application;            
            //app.SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(Application_SheetActivate);
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
            decorator = new SheetDecorator(messenger, sheet, syssheet);
            validator = new DataValidator(sheet, messenger, decorator);
            extractor = new DataExtractor(sheet, validator);
            determiner = new ValueDeterminer(sheet, messenger, validator);
            eventmanager = new EventManagement(sheet, messenger, determiner);
            
            Validate.Enabled = true;
            UploadSheet.Enabled = true;
            // GetTemplate.Enabled = false;

            decorator.SetDropDowns();
            decorator.GenerateHeader();

            eventmanager.SetEventHandlers();
        }

        private void Validate_Click(object sender, RibbonControlEventArgs e)
        {
            NotificationService notify = new NotificationService();

            Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = sheet.get_Range("A1", last);
            int lastUsedRow = last.Row;

            List<int> rows = new List<int>();
            for(int i = 2; i <= lastUsedRow; ++i)
            {
                rows.Add(i);
            }
            
            List<Handover> handoverlist = extractor.ExtractHandovers(rows);
            notify.ValidationComplete();

            if (handoverlist == null)
            {
                return;
            }
            
            validator.ValidateHandovers(handoverlist);
            notify.ValidationComplete();
        }
    }
}
