using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using MyntraExcelAddin.Entity;
using MyntraExcelAddin.Constant;
using MyntraExcelAddin.SystemObjects;

namespace MyntraExcelAddin.Service
{
    class DataExtractor
    {
        Excel._Worksheet sheet;
        DataValidator validator;

        public DataExtractor(Excel._Worksheet sheet, DataValidator validator)
        {
            this.sheet = sheet;
            this.validator = validator;
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

            /*********************** Fabric Extraction logic BEGINS ***********************/
            List<Fabric> tempfabriclist = new List<Fabric>();
            
            Fabric f1 = new Fabric();
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric1_quality)) {
                f1.quality = sheet.Cells[rowindex, ColumnNumber.fabric1_quality].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric1_impression)) {
                f1.impression = sheet.Cells[rowindex, ColumnNumber.fabric1_impression].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric1_baseColor)) {
                f1.baseColor = sheet.Cells[rowindex, ColumnNumber.fabric1_baseColor].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric1_printCode)) {
                f1.printCode = sheet.Cells[rowindex, ColumnNumber.fabric1_printCode].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric1_fpt)) {
                f1.fpt = sheet.Cells[rowindex, ColumnNumber.fabric1_fpt].Value;
            }
            tempfabriclist.Add(f1);

            Fabric f2 = new Fabric();
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric2_quality)) {
                f2.quality = sheet.Cells[rowindex, ColumnNumber.fabric2_quality].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric2_impression)) {
                f2.impression = sheet.Cells[rowindex, ColumnNumber.fabric2_impression].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric2_baseColor)) {
                f2.baseColor = sheet.Cells[rowindex, ColumnNumber.fabric2_baseColor].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric2_printCode)) {
                f2.printCode = sheet.Cells[rowindex, ColumnNumber.fabric2_printCode].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric2_fpt)) {
                f2.fpt = sheet.Cells[rowindex, ColumnNumber.fabric2_fpt].Value;
            }
            tempfabriclist.Add(f2);

            Fabric f3 = new Fabric();
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric3_quality)) {
                f3.quality = sheet.Cells[rowindex, ColumnNumber.fabric3_quality].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric3_impression)) {
                f3.impression = sheet.Cells[rowindex, ColumnNumber.fabric3_impression].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric3_baseColor)) {
                f3.baseColor = sheet.Cells[rowindex, ColumnNumber.fabric3_baseColor].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric3_printCode)) {
                f3.printCode = sheet.Cells[rowindex, ColumnNumber.fabric3_printCode].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric3_fpt)) {
                f3.fpt = sheet.Cells[rowindex, ColumnNumber.fabric3_fpt].Value;
            }
            tempfabriclist.Add(f3);

            Fabric f4 = new Fabric();
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric4_quality)) {
                f4.quality = sheet.Cells[rowindex, ColumnNumber.fabric4_quality].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric4_impression)) {
                f4.impression = sheet.Cells[rowindex, ColumnNumber.fabric4_impression].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric4_baseColor)) {
                f4.baseColor = sheet.Cells[rowindex, ColumnNumber.fabric4_baseColor].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric4_printCode)) {
                f4.printCode = sheet.Cells[rowindex, ColumnNumber.fabric4_printCode].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric4_fpt)) {
                f4.fpt = sheet.Cells[rowindex, ColumnNumber.fabric4_fpt].Value;
            }
            tempfabriclist.Add(f4);

            Fabric f5 = new Fabric();
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric5_quality)) {
                f5.quality = sheet.Cells[rowindex, ColumnNumber.fabric5_quality].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric5_impression)) {
                f5.impression = sheet.Cells[rowindex, ColumnNumber.fabric5_impression].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric5_baseColor)) {
                f5.baseColor = sheet.Cells[rowindex, ColumnNumber.fabric5_baseColor].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric5_printCode)) {
                f5.printCode = sheet.Cells[rowindex, ColumnNumber.fabric5_printCode].Value;
            }
            if (!validator.IsEmptyCell(rowindex, ColumnNumber.fabric5_fpt)) {
                f5.fpt = sheet.Cells[rowindex, ColumnNumber.fabric5_fpt].Value;
            }
            tempfabriclist.Add(f5);

            handover.fabrics = tempfabriclist;
            /*********************** Fabric Extraction logic ENDS ***********************/

            handover.dropName = sheet.Cells[rowindex, ColumnNumber.dropName].Value;
            handover.mrpRange = sheet.Cells[rowindex, ColumnNumber.mrpRange].Value;
            //handover.bmTarget = sheet.Cells[rowindex, ColumnNumber.bmTarget].Value;
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
            NotificationService notifier = new NotificationService();

            List<Handover> handoverList = new List<Handover>();
            foreach(int row in rows) {
                if(validator.HasEmptyCells(row)) {
                    return null;
                }

                Handover handover = ExtractHandoverFromRow(row);
                handoverList.Add(handover);
            }

            foreach(Handover h in handoverList)
            {
                System.Diagnostics.Debug.WriteLine(" ******************** Handover ******************** ");
                System.Diagnostics.Debug.WriteLine("handover.vanId = " + h.vanId);
                System.Diagnostics.Debug.WriteLine("handover.brand = " + h.brand);
                System.Diagnostics.Debug.WriteLine("handover.articleType = " + h.articleType);
                System.Diagnostics.Debug.WriteLine("handover.gender = " + h.gender);
                System.Diagnostics.Debug.WriteLine("handover.quantity = " + h.quantity);
                System.Diagnostics.Debug.WriteLine("handover.cluster = " + h.cluster);
                System.Diagnostics.Debug.WriteLine("handover.subcategory = " + h.subcategory);
                
                System.Diagnostics.Debug.WriteLine(" ******** Fabrics ******** ");
                foreach(Fabric f in h.fabrics) {
                    System.Diagnostics.Debug.WriteLine("quality = " + f.quality);
                    System.Diagnostics.Debug.WriteLine("impression = " + f.impression);
                    System.Diagnostics.Debug.WriteLine("baseColor = " + f.baseColor);
                    System.Diagnostics.Debug.WriteLine("printCode = " + f.printCode);
                    System.Diagnostics.Debug.WriteLine("fpt = " + f.fpt);
                    System.Diagnostics.Debug.WriteLine(" -------------------------- ");
                }

                System.Diagnostics.Debug.WriteLine("handover.dropName = " + h.dropName);
                System.Diagnostics.Debug.WriteLine("handover.mrpRange = " + h.mrpRange);
                System.Diagnostics.Debug.WriteLine("handover.bmTarget = " + h.bmTarget);
                System.Diagnostics.Debug.WriteLine("handover.sizeType = " + h.sizeType);
                System.Diagnostics.Debug.WriteLine("handover.bodyCode = " + h.bodyCode);
                System.Diagnostics.Debug.WriteLine("handover.dataSource = " + h.dataSource);
                System.Diagnostics.Debug.WriteLine("handover.dataSourceDetails = " + h.dataSourceDetails);
                System.Diagnostics.Debug.WriteLine("handover.color = " + h.color);
                System.Diagnostics.Debug.WriteLine("handover.isWashReferenced = " + h.isWashReferenced);
                System.Diagnostics.Debug.WriteLine("handover.pdpCatalogCallouts = " + h.pdpCatalogCallouts);
                System.Diagnostics.Debug.WriteLine("handover.source = " + h.source);

            }

            return handoverList;
        }
    }
}
