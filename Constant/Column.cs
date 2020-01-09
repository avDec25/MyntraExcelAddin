using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace MyntraExcelAddin.Constant
{
    static class ColumnMeta
    {
        public const int TotalColumns = 45;
        public static readonly IList<String> ColumnHeader = new ReadOnlyCollection<string> (new List<String> {                
                "Repeated? (true/false)",
                "Style Id (if repeated)",
                "Van Id",
                "Brand",
                "Gender",
                "Article Type",
                "Quantity",
                "Cluster",
                "Sub-Category",
                "Fabric 1- Quality",
                "Fabric 1- solid / print",
                "Fabric 1- Base Color (Pantone)",
                "Fabric 1- Print Code",
                "Fabric 1- Technique",
                "Fabric 2- Quality",
                "Fabric 2- solid / print",
                "Fabric 2- Base Color (Pantone)",
                "Fabric 2- Print Code",
                "Fabric 2- Technique",
                "Fabric 3- Quality",
                "Fabric 3- solid / print",
                "Fabric 3- Base Color (Pantone)",
                "Fabric 3- Print Code",
                "Fabric 3- Technique",
                "Fabric 4- Quality",
                "Fabric 4- solid / print",
                "Fabric 4- Base Color (Pantone)",
                "Fabric 4- Print Code",
                "Fabric 4- Technique",
                "Fabric 5- Quality",
                "Fabric 5- solid / print",
                "Fabric 5- Base Color (Pantone)",
                "Fabric 5- Print Code",
                "Fabric 5- Technique",
                "Drop Name",
                "MRP Range",
                "BM Target",
                "Size Type",
                "Body Code",
                "Data Source",
                "Data Source Details",
                "Color (will be mentioned in Catalog)",
                "Wash Referenced? (If any) (true/false)",
                "Catalog PDP Callouts (if any)",
                "Source",
                "Handover ID"
        });
    }

    static class Header
    {
        public static readonly IList<String> Name = new ReadOnlyCollection<string>(new List<String> {
            "",
            "Repeated",
            "Style Id",
            "Van Id",
            "Brand",
            "Gender",
            "Article Type",
            "Quantity",
            "Cluster",
            "Sub-Category",
            "Fabric 1- Quality",
            "Fabric 1- solid / print",
            "Fabric 1- Base Color",
            "Fabric 1- Print Code",
            "Fabric 1- Technique",
            "Fabric 2- Quality",
            "Fabric 2- solid / print",
            "Fabric 2- Base Color",
            "Fabric 2- Print Code",
            "Fabric 2- Technique",
            "Fabric 3- Quality",
            "Fabric 3- solid / print",
            "Fabric 3- Base Color",
            "Fabric 3- Print Code",
            "Fabric 3- Technique",
            "Fabric 4- Quality",
            "Fabric 4- solid / print",
            "Fabric 4- Base Color",
            "Fabric 4- Print Code",
            "Fabric 4- Technique",
            "Fabric 5- Quality",
            "Fabric 5- solid / print",
            "Fabric 5- Base Color",
            "Fabric 5- Print Code",
            "Fabric 5- Technique",
            "Drop Name",
            "MRP Range",
            "BM Target",
            "Size Type",
            "Body Code",
            "Data Source",
            "Data Source Details",
            "Color",
            "Wash Referenced?",
            "Catalog PDP Callouts",
            "Source",
            "Handover ID"
        });
    }

    static class ColumnName
    {
        public const string repeated = "A";
        public const string styleid = "B";
        public const string vanId = "C";
        public const string brand = "D";
        public const string gender = "E";
        public const string articleType = "F";
        public const string quantity = "G";
        public const string cluster = "H";
        public const string subcategory = "I";
        public const string fabric1_quality = "J";
        public const string fabric1_impression = "K";
        public const string fabric1_baseColor = "L";
        public const string fabric1_printCode = "M";
        public const string fabric1_fpt = "N";
        public const string fabric2_quality = "O";
        public const string fabric2_impression = "P";
        public const string fabric2_baseColor = "Q";
        public const string fabric2_printCode = "R";
        public const string fabric2_fpt = "S";
        public const string fabric3_quality = "T";
        public const string fabric3_impression = "U";
        public const string fabric3_baseColor = "V";
        public const string fabric3_printCode = "W";
        public const string fabric3_fpt = "X";
        public const string fabric4_quality = "Y";
        public const string fabric4_impression = "Z";
        public const string fabric4_baseColor = "AA";
        public const string fabric4_printCode = "AB";
        public const string fabric4_fpt = "AC";
        public const string fabric5_quality = "AD";
        public const string fabric5_impression = "AE";
        public const string fabric5_baseColor = "AF";
        public const string fabric5_printCode = "AG";
        public const string fabric5_fpt = "AH";
        public const string dropName = "AI";
        public const string mrpRange = "AJ";
        public const string bmTarget = "AK";
        public const string sizeType = "AL";
        public const string bodyCode = "AM";
        public const string dataSource = "AN";
        public const string dataSourceDetails = "AO";
        public const string color = "AP";
        public const string isWashReferenced = "AQ";
        public const string pdpCatalogCallouts = "AR";
        public const string source = "AS";
        public const string handoverId = "AT";        
    }

    static class ColumnNumber
    {   
        public const int repeated = 1;
        public const int styleid = 2;
        public const int vanId = 3;
        public const int brand = 4;
        public const int gender = 5;
        public const int articleType = 6;
        public const int quantity = 7;
        public const int cluster = 8;
        public const int subcategory = 9;
        public const int fabric1_quality = 10;
        public const int fabric1_impression = 11;
        public const int fabric1_baseColor = 12;
        public const int fabric1_printCode = 13;
        public const int fabric1_fpt = 14;
        public const int fabric2_quality = 15;
        public const int fabric2_impression = 16;
        public const int fabric2_baseColor = 17;
        public const int fabric2_printCode = 18;
        public const int fabric2_fpt = 19;
        public const int fabric3_quality = 20;
        public const int fabric3_impression = 21;
        public const int fabric3_baseColor = 22;
        public const int fabric3_printCode = 23;
        public const int fabric3_fpt = 24;
        public const int fabric4_quality = 25;
        public const int fabric4_impression = 26;
        public const int fabric4_baseColor = 27;
        public const int fabric4_printCode = 28;
        public const int fabric4_fpt = 29;
        public const int fabric5_quality = 30;
        public const int fabric5_impression = 31;
        public const int fabric5_baseColor = 32;
        public const int fabric5_printCode = 33;
        public const int fabric5_fpt = 34;
        public const int dropName = 35;
        public const int mrpRange = 36;
        public const int bmTarget = 37;
        public const int sizeType = 38;
        public const int bodyCode = 39;
        public const int dataSource = 40;
        public const int dataSourceDetails = 41;
        public const int color = 42;
        public const int isWashReferenced = 43;
        public const int pdpCatalogCallouts = 44;
        public const int source = 45;
        public const int handoverId = 46;
    }
}
