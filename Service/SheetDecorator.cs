using MyntraExcelAddin.Constant;
using MyntraExcelAddin.SystemObjects;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace MyntraExcelAddin.Service
{
    public class SheetDecorator
    {
        public Excel._Worksheet sheet;
        public Excel._Worksheet syssheet;
        public ExternalServiceMessenger messenger;

        public SheetDecorator(ExternalServiceMessenger msngr, Excel._Worksheet xlsheet, Excel._Worksheet systemsheet)
        {
            messenger = msngr;
            sheet = xlsheet;
            syssheet = systemsheet;
        }

        public void HighlightErrorAtCell(int row, int col, string message)
        {
            sheet.Cells[row, col].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 148, 148));                        
            sheet.Cells[row, col].Validation.InputMessage = message;
        }

        public void ClearAllErrors(int row)
        {
            sheet.Cells[row, ColumnNumber.sizeType].Validation.InputMessage = "";
            sheet.Cells[row, ColumnNumber.sizeType].ClearFormats();

            sheet.Cells[row, ColumnNumber.brand].Validation.InputMessage = "";
            sheet.Cells[row, ColumnNumber.brand].ClearFormats();

            sheet.Cells[row, ColumnNumber.articleType].Validation.InputMessage = "";
            sheet.Cells[row, ColumnNumber.articleType].ClearFormats();

            sheet.Cells[row, ColumnNumber.gender].Validation.InputMessage = "";
            sheet.Cells[row, ColumnNumber.gender].ClearFormats();

            sheet.Cells[row, ColumnNumber.quantity].Validation.InputMessage = "";
            sheet.Cells[row, ColumnNumber.quantity].ClearFormats();

            sheet.Cells[row, ColumnNumber.cluster].Validation.InputMessage = "";
            sheet.Cells[row, ColumnNumber.cluster].ClearFormats();

            sheet.Cells[row, ColumnNumber.subcategory].Validation.InputMessage = "";
            sheet.Cells[row, ColumnNumber.subcategory].ClearFormats();

            sheet.Cells[row, ColumnNumber.bmTarget].Validation.InputMessage = "";
            sheet.Cells[row, ColumnNumber.bmTarget].ClearFormats();

            sheet.Cells[row, ColumnNumber.bodyCode].Validation.InputMessage = "";
            sheet.Cells[row, ColumnNumber.bodyCode].ClearFormats();

            sheet.Cells[row, ColumnNumber.color].Validation.InputMessage = "";
            sheet.Cells[row, ColumnNumber.color].ClearFormats();
        }

        public void GenerateHeader()
        {
            Excel.Range headerRange = sheet.Range["A1", "AT1"];
            headerRange.ClearFormats();
            headerRange.Validation.Delete();

            headerRange.WrapText = true;
            headerRange.Font.Name = "Arial";
            headerRange.Font.Size = 16;
            headerRange.ColumnWidth = 18;
            headerRange.RowHeight = 88;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            headerRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            sheet.Columns["AU:XFD"].EntireColumn.Hidden = true;

            sheet.Range["A1", "I1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 0));
            sheet.Range["J1", "N1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(221, 235, 247));
            sheet.Range["O1", "S1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189, 215, 238));
            sheet.Range["T1", "X1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(221, 235, 247));
            sheet.Range["Y1", "AC1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189, 215, 238));
            sheet.Range["AD1", "AH1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(221, 235, 247));
            sheet.Range["AI1", "AL1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 0));

            sheet.Rows["1:1"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;

            int col = 1;
            foreach (string headerCol in ColumnMeta.ColumnHeader)
            {
                sheet.Cells[1, col] = headerCol;
                ++col;
            }
        }

        public void AddFakeValidations()
        {
            for (int ci = ColumnNumber.repeated; ci <= ColumnNumber.handoverId; ++ci)
            {
                // In case, the input message is to be displayed on a cell which previously has no validations enabled; 
                // First Enable the validations:
                if (ci == ColumnNumber.fabric1_printCode ||
                    ci == ColumnNumber.fabric1_quality ||
                    ci == ColumnNumber.fabric1_baseColor ||
                    ci == ColumnNumber.fabric2_printCode ||
                    ci == ColumnNumber.fabric2_quality ||
                    ci == ColumnNumber.fabric2_baseColor ||
                    ci == ColumnNumber.fabric3_printCode ||
                    ci == ColumnNumber.fabric3_quality ||
                    ci == ColumnNumber.fabric3_baseColor ||
                    ci == ColumnNumber.fabric4_printCode ||
                    ci == ColumnNumber.fabric4_quality ||
                    ci == ColumnNumber.fabric4_baseColor ||
                    ci == ColumnNumber.fabric5_printCode ||
                    ci == ColumnNumber.fabric5_quality ||
                    ci == ColumnNumber.fabric5_baseColor ||
                    ci == ColumnNumber.styleid ||
                    ci == ColumnNumber.vanId ||
                    ci == ColumnNumber.quantity ||
                    ci == ColumnNumber.dropName ||
                    ci == ColumnNumber.mrpRange ||
                    ci == ColumnNumber.bmTarget ||
                    ci == ColumnNumber.dataSourceDetails ||
                    ci == ColumnNumber.isWashReferenced ||
                    ci == ColumnNumber.pdpCatalogCallouts ||
                    ci == ColumnNumber.handoverId)
                {
                    sheet.Columns[ci].Validation.Add(Excel.XlDVType.xlValidateInputOnly, Excel.XlDVAlertStyle.xlValidAlertInformation,
                        Excel.XlFormatConditionOperator.xlBetween, Type.Missing, Type.Missing);
                }
            }
        }

        public void SetDropDowns()
        {
            DropDownData ddd = messenger.GetDropDownData("repeated,brand,impression,articletype,gender,bodycode,cluster,color,subcategory,fpt,sizetype,datasource,source");
            if (ddd == null)
            {
                return;
            }

            putDropDownData(ddd.repeated, ColumnName.repeated, ColumnNumber.repeated);
            putDropDownData(ddd.brand, ColumnName.brand, ColumnNumber.brand);
            putDropDownData(ddd.gender, ColumnName.gender, ColumnNumber.gender);
            putDropDownData(ddd.articletype, ColumnName.articleType, ColumnNumber.articleType);
            putDropDownData(ddd.cluster, ColumnName.cluster, ColumnNumber.cluster);
            putDropDownData(ddd.subcategory, ColumnName.subcategory, ColumnNumber.subcategory);

            putDropDownData(ddd.impression, ColumnName.fabric1_impression, ColumnNumber.fabric1_impression);
            putDropDownData(ddd.fpt, ColumnName.fabric1_fpt, ColumnNumber.fabric1_fpt);

            putDropDownData(ddd.impression, ColumnName.fabric2_impression, ColumnNumber.fabric2_impression);
            putDropDownData(ddd.fpt, ColumnName.fabric2_fpt, ColumnNumber.fabric2_fpt);

            putDropDownData(ddd.impression, ColumnName.fabric3_impression, ColumnNumber.fabric3_impression);
            putDropDownData(ddd.fpt, ColumnName.fabric3_fpt, ColumnNumber.fabric3_fpt);

            putDropDownData(ddd.impression, ColumnName.fabric4_impression, ColumnNumber.fabric4_impression);
            putDropDownData(ddd.fpt, ColumnName.fabric4_fpt, ColumnNumber.fabric4_fpt);

            putDropDownData(ddd.impression, ColumnName.fabric5_impression, ColumnNumber.fabric5_impression);
            putDropDownData(ddd.fpt, ColumnName.fabric5_fpt, ColumnNumber.fabric5_fpt);

            putDropDownData(ddd.sizetype, ColumnName.sizeType, ColumnNumber.sizeType);
            putDropDownData(ddd.bodycode, ColumnName.bodyCode, ColumnNumber.bodyCode);
            putDropDownData(ddd.datasource, ColumnName.dataSource, ColumnNumber.dataSource);
            putDropDownData(ddd.color, ColumnName.color, ColumnNumber.color);
            putDropDownData(ddd.source, ColumnName.source, ColumnNumber.source);
        }

        private void putDropDownData(string[] data, string colname, int colnum)
        {
            int i = 2;
            Array.Sort(data);
            foreach (string text in data)
            {
                syssheet.Cells[i, colnum] = text;
                i++;
            }
            --i;
            sheet.Columns[colnum].Validation.Add(
                Excel.XlDVType.xlValidateList,
                Excel.XlDVAlertStyle.xlValidAlertStop,
                Type.Missing,
                "=" + syssheet.Name + "!$" + colname + "$2:$" + colname + "$" + i);
        }
    }
}
