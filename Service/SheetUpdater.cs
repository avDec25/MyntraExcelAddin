using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using MyntraExcelAddin.Constant;
using MyntraExcelAddin.Entity;

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

        public void ClearSheetData()
        {
            Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = sheet.get_Range("A2", last);            

            range.Rows.ClearContents();
        }

        public void PutDownloadedHandoversOnSheet(List<Handover> handoverlist)
        {
            ClearSheetData();
            for(int i = 0; i < handoverlist.Count; ++i)
            {
                sheet.Cells[i + 2, ColumnName.repeated].Value = handoverlist[i].repeated;
                sheet.Cells[i + 2, ColumnName.vanId].Value = handoverlist[i].vanId;
                sheet.Cells[i + 2, ColumnName.brand].Value = handoverlist[i].brand;
                sheet.Cells[i + 2, ColumnName.articleType].Value = handoverlist[i].articleType;
                sheet.Cells[i + 2, ColumnName.gender].Value = handoverlist[i].gender;
                sheet.Cells[i + 2, ColumnName.quantity].Value = handoverlist[i].quantity;
                sheet.Cells[i + 2, ColumnName.cluster].Value = handoverlist[i].cluster;
                sheet.Cells[i + 2, ColumnName.subcategory].Value = handoverlist[i].subcategory;

                sheet.Cells[i + 2, ColumnName.fabric1_quality].Value = handoverlist[i].fabrics[0].quality;
                sheet.Cells[i + 2, ColumnName.fabric1_impression].Value = handoverlist[i].fabrics[0].impression;
                sheet.Cells[i + 2, ColumnName.fabric1_baseColor].Value = handoverlist[i].fabrics[0].baseColor;
                sheet.Cells[i + 2, ColumnName.fabric1_printCode].Value = handoverlist[i].fabrics[0].printCode;
                sheet.Cells[i + 2, ColumnName.fabric1_fpt].Value = handoverlist[i].fabrics[0].fpt;

                if (handoverlist[i].fabrics.Count > 1)
                {
                    sheet.Cells[i + 2, ColumnName.fabric2_quality].Value = handoverlist[i].fabrics[1].quality;
                    sheet.Cells[i + 2, ColumnName.fabric2_impression].Value = handoverlist[i].fabrics[1].impression;
                    sheet.Cells[i + 2, ColumnName.fabric2_baseColor].Value = handoverlist[i].fabrics[1].baseColor;
                    sheet.Cells[i + 2, ColumnName.fabric2_printCode].Value = handoverlist[i].fabrics[1].printCode;
                    sheet.Cells[i + 2, ColumnName.fabric2_fpt].Value = handoverlist[i].fabrics[1].fpt;
                }

                if (handoverlist[i].fabrics.Count > 2)
                {
                    sheet.Cells[i + 2, ColumnName.fabric3_quality].Value = handoverlist[i].fabrics[2].quality;
                    sheet.Cells[i + 2, ColumnName.fabric3_impression].Value = handoverlist[i].fabrics[2].impression;
                    sheet.Cells[i + 2, ColumnName.fabric3_baseColor].Value = handoverlist[i].fabrics[2].baseColor;
                    sheet.Cells[i + 2, ColumnName.fabric3_printCode].Value = handoverlist[i].fabrics[2].printCode;
                    sheet.Cells[i + 2, ColumnName.fabric3_fpt].Value = handoverlist[i].fabrics[2].fpt;
                }

                if(handoverlist[i].fabrics.Count > 3) {
                    sheet.Cells[i + 2, ColumnName.fabric4_quality].Value = handoverlist[i].fabrics[3].quality;
                    sheet.Cells[i + 2, ColumnName.fabric4_impression].Value = handoverlist[i].fabrics[3].impression;
                    sheet.Cells[i + 2, ColumnName.fabric4_baseColor].Value = handoverlist[i].fabrics[3].baseColor;
                    sheet.Cells[i + 2, ColumnName.fabric4_printCode].Value = handoverlist[i].fabrics[3].printCode;
                    sheet.Cells[i + 2, ColumnName.fabric4_fpt].Value = handoverlist[i].fabrics[3].fpt;
                }

                if(handoverlist[i].fabrics.Count > 4) {
                    sheet.Cells[i + 2, ColumnName.fabric5_quality].Value = handoverlist[i].fabrics[4].quality;
                    sheet.Cells[i + 2, ColumnName.fabric5_impression].Value = handoverlist[i].fabrics[4].impression;
                    sheet.Cells[i + 2, ColumnName.fabric5_baseColor].Value = handoverlist[i].fabrics[4].baseColor;
                    sheet.Cells[i + 2, ColumnName.fabric5_printCode].Value = handoverlist[i].fabrics[4].printCode;
                    sheet.Cells[i + 2, ColumnName.fabric5_fpt].Value = handoverlist[i].fabrics[4].fpt;
                }

                sheet.Cells[i + 2, ColumnName.dropName].Value = handoverlist[i].dropName;
                sheet.Cells[i + 2, ColumnName.mrpRange].Value = handoverlist[i].mrpRangeLower + "-" + handoverlist[i].mrpRangeUpper;
                sheet.Cells[i + 2, ColumnName.bmTarget].Value = handoverlist[i].bmTarget;
                sheet.Cells[i + 2, ColumnName.sizeType].Value = handoverlist[i].sizeType;
                sheet.Cells[i + 2, ColumnName.bodyCode].Value = handoverlist[i].bodyCode;
                sheet.Cells[i + 2, ColumnName.dataSource].Value = handoverlist[i].dataSource;
                sheet.Cells[i + 2, ColumnName.dataSourceDetails].Value = handoverlist[i].dataSourceDetails;
                sheet.Cells[i + 2, ColumnName.color].Value = handoverlist[i].color;
                sheet.Cells[i + 2, ColumnName.isWashReferenced].Value = handoverlist[i].isWashReferenced;
                sheet.Cells[i + 2, ColumnName.pdpCatalogCallouts].Value = handoverlist[i].pdpCatalogCallouts;
                sheet.Cells[i + 2, ColumnName.source].Value = handoverlist[i].source;
                sheet.Cells[i + 2, ColumnName.handoverId].Value = handoverlist[i].id;
            }
        }

    }
}
