using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using MyntraExcelAddin.Service;
using MyntraExcelAddin.SystemObjects;
using MyntraExcelAddin.Entity;

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
        SheetUpdater sheetUpdater;

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
            decorator = new SheetDecorator(messenger, sheet, syssheet);
            validator = new DataValidator(sheet, messenger, decorator);
            extractor = new DataExtractor(sheet, validator);
            determiner = new ValueDeterminer(sheet, messenger, validator);
            eventmanager = new EventManagement(sheet, messenger, determiner);
            sheetUpdater = new SheetUpdater(sheet);


            Validate.Enabled = true;
            UploadSheet.Enabled = true;
            UpdateSheet.Enabled = false;
            // GetTemplate.Enabled = false;

            decorator.SetDropDowns();
            decorator.GenerateHeader();
            decorator.AddFakeValidations(); // 1 time required to enable Validations.InputMessage box under a cell.

            eventmanager.SetEventHandlers();
        }

        private List<Handover> GetHandoverList()
        {
            Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = sheet.get_Range("A1", last);
            int lastUsedRow = last.Row;

            List<int> rows = new List<int>();
            for (int i = 2; i <= lastUsedRow; ++i)
            {
                rows.Add(i);
            }

            return extractor.ExtractHandovers(rows);
        }

        private void Validate_Click(object sender, RibbonControlEventArgs e)
        {
            NotificationService notify = new NotificationService();
            List<Handover> handoverlist = GetHandoverList();

            if (handoverlist == null)
            {
                notify.ProcessComplete("Validation Service", "failed");
                return;
            }

            if (validator.ValidateHandovers(handoverlist)) 
            {
                notify.ProcessComplete("Validation Service", "success");
            }
            else
            {
                notify.ProcessComplete("Validation Service", "failed");
            }
        }

        private void UploadSheet_Click(object sender, RibbonControlEventArgs e)
        {
            NotificationService notify = new NotificationService();
            List<Handover> handoverlist = GetHandoverList();

            if (handoverlist == null)
            {
                notify.ProcessComplete("Upload Service", "failed");
                return;
            }

            List<long> savedHandoverIds = messenger.SubmitHandovers(handoverlist);            
            if (savedHandoverIds.Count > 0)
            {
                sheetUpdater.HandoverIdsUpdate(savedHandoverIds);
                UpdateSheet.Enabled = true;
                notify.ProcessComplete("Upload Service", "success");                
            }
            else
            {
                notify.ProcessComplete("Upload Service", "failed");
            }
        }

        private void UpdateSheet_Click(object sender, RibbonControlEventArgs e)
        {
            NotificationService notify = new NotificationService();
            List<Handover> handoverlist = GetHandoverList();

            if (validator.ValidateHandovers(handoverlist))
            {
                messenger.UpdateHandovers(handoverlist);
                notify.ProcessComplete("Update Handovers", "success");
            }
            else
            {
                notify.ProcessComplete("Update Handovers", "failed");
            }
        }
    }
}
