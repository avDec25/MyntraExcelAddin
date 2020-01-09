using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using MyntraExcelAddin.Entity;
using MyntraExcelAddin.Constant;

namespace MyntraExcelAddin.Service
{
    class DataExtractor
    {
        Excel._Worksheet sheet;

        public DataExtractor(Excel._Worksheet sheet)
        {
            this.sheet = sheet;
        }

        private Handover ExtractHandoverFromRow(int rowindex)
        {
            Handover handover = new Handover();
            handover.vanId = sheet.Cells[rowindex, ColumnNumber.vanId].Value;
            handover.brand = sheet.Cells[rowindex, ColumnNumber.brand].Value;
            handover.articleType = sheet.Cells[rowindex, ColumnNumber.articleType].Value;
            handover.gender = sheet.Cells[rowindex, ColumnNumber.gender].Value;
            handover.quantity = sheet.Cells[rowindex, ColumnNumber.quantity].Value;
            handover.cluster = sheet.Cells[rowindex, ColumnNumber.cluster].Value;
            handover.subcategory = sheet.Cells[rowindex, ColumnNumber.subcategory].Value;
            
            //handover.fabrics = sheet.Cells[rowindex, ColumnNumber.fabrics].Value;

            handover.dropName = sheet.Cells[rowindex, ColumnNumber.dropName].Value;
            handover.mrpRange = sheet.Cells[rowindex, ColumnNumber.mrpRange].Value;
            handover.bmTarget = sheet.Cells[rowindex, ColumnNumber.bmTarget].Value;
            handover.sizeType = sheet.Cells[rowindex, ColumnNumber.sizeType].Value;
            handover.bodyCode = sheet.Cells[rowindex, ColumnNumber.bodyCode].Value;
            handover.dataSource = sheet.Cells[rowindex, ColumnNumber.dataSource].Value;
            handover.dataSourceDetails = sheet.Cells[rowindex, ColumnNumber.dataSourceDetails].Value;
            handover.color = sheet.Cells[rowindex, ColumnNumber.color].Value;
            handover.isWashReferenced = sheet.Cells[rowindex, ColumnNumber.isWashReferenced].Value;
            handover.pdpCatalogCallouts = sheet.Cells[rowindex, ColumnNumber.pdpCatalogCallouts].Value;
            handover.source = sheet.Cells[rowindex, ColumnNumber.source].Value;


            return handover;
        }

        public List<Handover> ExtractHandovers(List<int> rows)
        {
            DataValidator validator = new DataValidator(sheet);
            NotificationService notifier = new NotificationService();

            List<Handover> handoverList = new List<Handover>();
            foreach(int row in rows) {
                if(validator.HasEmptyCells(row)) {
                    return null;
                }

                Handover handover = ExtractHandoverFromRow(row);
            }
            return handoverList;
        }
    }
}
