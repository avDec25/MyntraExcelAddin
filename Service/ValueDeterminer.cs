using MyntraExcelAddin.SystemObjects;
using MyntraExcelAddin.Constant;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace MyntraExcelAddin.Service
{
    public class ValueDeterminer
    {
        public Excel._Worksheet sheet;
        ExternalServiceMessenger messenger;
        DataValidator validator;

        public ValueDeterminer(Excel._Worksheet sheet, ExternalServiceMessenger messenger, DataValidator validator)
        {
            this.sheet = sheet;
            this.messenger = messenger;
            this.validator = validator;
        }

        public Double DetermineBmTarget(int row)
        {
            string brand;
            string articletype;
            string gender;
            bool repeated;

            if(validator.IsEmptyCell(row, ColumnNumber.brand) ||
                validator.IsEmptyCell(row, ColumnNumber.articleType) ||
                validator.IsEmptyCell(row, ColumnNumber.gender) ||
                validator.IsEmptyCell(row, ColumnNumber.repeated))
            {
                return -1.0;
            } 
            else
            {
                brand = sheet.Cells[row, ColumnName.brand].Value;
                articletype = sheet.Cells[row, ColumnName.articleType].Value;
                gender = sheet.Cells[row, ColumnName.gender].Value;
                repeated = sheet.Cells[row, ColumnName.repeated].Value;
            }
            return messenger.RetrieveBMTargetValue(brand, articletype, gender, repeated);
        }
    }
}
