using MyntraExcelAddin.SystemObjects;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using MyntraExcelAddin.Constant;

namespace MyntraExcelAddin.Service
{
    class EventManagement
    {
        Excel._Worksheet sheet;
        public ExternalServiceMessenger messenger;
        ValueDeterminer determiner;

        public EventManagement(Excel._Worksheet sheet, ExternalServiceMessenger messenger, ValueDeterminer determiner)
        {
            this.sheet = sheet;
            this.messenger = messenger;
            this.determiner = determiner;
        }

        public void SetEventHandlers()
        {
            //sheet.UsedRange.Columns["F:F", Type.Missing]
            Globals.ThisAddIn.Application.SheetChange += new Excel.AppEvents_SheetChangeEventHandler(SheetChange);
        }

        private void SheetChange(object Sh, Excel.Range Target)
        {
            switch(Target.Column) {                
                case ColumnNumber.repeated:
                    System.Diagnostics.Debug.WriteLine("Updated = repeated");
                    PossiblyDetermineBmTarget(Target.Row);
                    break;
                
                case ColumnNumber.styleid:
                    System.Diagnostics.Debug.WriteLine("Updated = styleid");
                    break;
                
                case ColumnNumber.vanId:
                    System.Diagnostics.Debug.WriteLine("Updated = vanId");
                    break;
                
                case ColumnNumber.brand:
                    System.Diagnostics.Debug.WriteLine("Updated = brand");
                    PossiblyDetermineBmTarget(Target.Row);
                    break;
                
                case ColumnNumber.gender:
                    System.Diagnostics.Debug.WriteLine("Updated = gender");
                    PossiblyDetermineBmTarget(Target.Row);
                    break;
                
                case ColumnNumber.articleType:
                    System.Diagnostics.Debug.WriteLine("Updated = articleType");
                    PossiblyDetermineBmTarget(Target.Row);
                    break;
                
                case ColumnNumber.quantity:
                    System.Diagnostics.Debug.WriteLine("Updated = quantity");
                    break;
                
                case ColumnNumber.cluster:
                    System.Diagnostics.Debug.WriteLine("Updated = cluster");
                    break;
                
                case ColumnNumber.subcategory:
                    System.Diagnostics.Debug.WriteLine("Updated = subcategory");
                    break;
                
                case ColumnNumber.fabric1_quality:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric1_quality");
                    break;
                
                case ColumnNumber.fabric1_impression:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric1_impression");
                    break;
                
                case ColumnNumber.fabric1_baseColor:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric1_baseColor");
                    break;
                
                case ColumnNumber.fabric1_printCode:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric1_printCode");
                    break;
                
                case ColumnNumber.fabric1_fpt:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric1_fpt");
                    break;
                
                case ColumnNumber.fabric2_quality:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric2_quality");
                    break;
                
                case ColumnNumber.fabric2_impression:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric2_impression");
                    break;
                
                case ColumnNumber.fabric2_baseColor:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric2_baseColor");
                    break;
                
                case ColumnNumber.fabric2_printCode:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric2_printCode");
                    break;
                
                case ColumnNumber.fabric2_fpt:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric2_fpt");
                    break;
                
                case ColumnNumber.fabric3_quality:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric3_quality");
                    break;
                
                case ColumnNumber.fabric3_impression:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric3_impression");
                    break;
                
                case ColumnNumber.fabric3_baseColor:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric3_baseColor");
                    break;
                
                case ColumnNumber.fabric3_printCode:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric3_printCode");
                    break;
                
                case ColumnNumber.fabric3_fpt:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric3_fpt");
                    break;
                
                case ColumnNumber.fabric4_quality:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric4_quality");
                    break;
                
                case ColumnNumber.fabric4_impression:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric4_impression");
                    break;
                
                case ColumnNumber.fabric4_baseColor:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric4_baseColor");
                    break;
                
                case ColumnNumber.fabric4_printCode:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric4_printCode");
                    break;
                
                case ColumnNumber.fabric4_fpt:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric4_fpt");
                    break;
                
                case ColumnNumber.fabric5_quality:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric5_quality");
                    break;
                
                case ColumnNumber.fabric5_impression:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric5_impression");
                    break;
                
                case ColumnNumber.fabric5_baseColor:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric5_baseColor");
                    break;
                
                case ColumnNumber.fabric5_printCode:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric5_printCode");
                    break;
                
                case ColumnNumber.fabric5_fpt:
                    System.Diagnostics.Debug.WriteLine("Updated = fabric5_fpt");
                    break;
                
                case ColumnNumber.dropName:
                    System.Diagnostics.Debug.WriteLine("Updated = dropName");
                    break;
                
                case ColumnNumber.mrpRange:
                    System.Diagnostics.Debug.WriteLine("Updated = mrpRange");
                    break;
                
                case ColumnNumber.bmTarget:
                    System.Diagnostics.Debug.WriteLine("Updated = bmTarget");
                    break;
                
                case ColumnNumber.sizeType:
                    System.Diagnostics.Debug.WriteLine("Updated = sizeType");
                    break;
                
                case ColumnNumber.bodyCode:
                    System.Diagnostics.Debug.WriteLine("Updated = bodyCode");
                    break;
                
                case ColumnNumber.dataSource:
                    System.Diagnostics.Debug.WriteLine("Updated = dataSource");
                    break;
                
                case ColumnNumber.dataSourceDetails:
                    System.Diagnostics.Debug.WriteLine("Updated = dataSourceDetails");
                    break;
                
                case ColumnNumber.color:
                    System.Diagnostics.Debug.WriteLine("Updated = color");
                    break;
                
                case ColumnNumber.isWashReferenced:
                    System.Diagnostics.Debug.WriteLine("Updated = isWashReferenced");
                    break;
                
                case ColumnNumber.pdpCatalogCallouts:
                    System.Diagnostics.Debug.WriteLine("Updated = pdpCatalogCallouts");
                    break;
                
                case ColumnNumber.source:
                    System.Diagnostics.Debug.WriteLine("Updated = source");
                    break;
                
                case ColumnNumber.handoverId:
                    System.Diagnostics.Debug.WriteLine("Updated = id");
                    break;
            }
            
        }

        private void PossiblyDetermineBmTarget(int row)
        {
            sheet.Cells[row, ColumnNumber.bmTarget].Value = determiner.DetermineBmTarget(row);
        }

        private string CellAddress(Excel.Range c)
        {
            return c.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
        }

    }

    // Reference:
    // http://blogs.infoextract.in/excel-events-using-c-dot-net/
    //System.Diagnostics.Debug.WriteLine(Target.Row);
    //System.Diagnostics.Debug.WriteLine(Target.Column);            
    //System.Diagnostics.Debug.WriteLine(Target.AddressLocal[false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);    
    //System.Diagnostics.Debug.WriteLine(Target.AddressLocal[true, true, Excel.XlReferenceStyle.xlA1, true, Type.Missing]);
}
