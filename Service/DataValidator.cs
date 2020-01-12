﻿using System;
using System.Collections.Generic;
using MyntraExcelAddin.Entity;
using MyntraExcelAddin.SystemObjects;
using MyntraExcelAddin.Constant;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace MyntraExcelAddin.Service
{
    class DataValidator
    {
        Excel._Worksheet sheet;
        public ExternalServiceMessenger messenger;
        public SheetDecorator decorator;

        public DataValidator(Excel._Worksheet sheet, ExternalServiceMessenger messenger, SheetDecorator decorator)
        {
            this.sheet = sheet;
            this.messenger = messenger;
            this.decorator = decorator;
        }

        public Boolean IsEmptyCell(int r, int c)
        {
            return (sheet.Cells[r, c] == null || sheet.Cells[r, c].Value2 == null || sheet.Cells[r, c].Value2.ToString() == "");
        }

        public Boolean HasEmptyCells(int row)
        {
            NotificationService notifier = new NotificationService();
            List<int> emptycols = new List<int>();

            if (IsEmptyCell(row, ColumnNumber.repeated)) {
                emptycols.Add(ColumnNumber.repeated);
            }
            //if (IsEmptyCell(row, ColumnNumber.styleid))
            //{
            //    emptycols.Add(ColumnNumber.styleid);
            //}
            if (IsEmptyCell(row, ColumnNumber.vanId)) {
                emptycols.Add(ColumnNumber.vanId);
            }
            if (IsEmptyCell(row, ColumnNumber.brand)) {
                emptycols.Add(ColumnNumber.brand);
            }
            if (IsEmptyCell(row, ColumnNumber.gender)) {
                emptycols.Add(ColumnNumber.gender);
            }
            if (IsEmptyCell(row, ColumnNumber.articleType)) {
                emptycols.Add(ColumnNumber.articleType);
            }
            if (IsEmptyCell(row, ColumnNumber.quantity)) {
                emptycols.Add(ColumnNumber.quantity);
            }
            if (IsEmptyCell(row, ColumnNumber.cluster)) {
                emptycols.Add(ColumnNumber.cluster);
            }
            if (IsEmptyCell(row, ColumnNumber.subcategory)) {
                emptycols.Add(ColumnNumber.subcategory);
            }
            if (IsEmptyCell(row, ColumnNumber.fabric1_quality)) {
                emptycols.Add(ColumnNumber.fabric1_quality);
            }
            if (IsEmptyCell(row, ColumnNumber.fabric1_impression)) {
                emptycols.Add(ColumnNumber.fabric1_impression);
            }
            if (IsEmptyCell(row, ColumnNumber.fabric1_baseColor)) {
                emptycols.Add(ColumnNumber.fabric1_baseColor);
            }
            if (IsEmptyCell(row, ColumnNumber.fabric1_printCode)) {
                emptycols.Add(ColumnNumber.fabric1_printCode);
            }
            if (IsEmptyCell(row, ColumnNumber.fabric1_fpt)) {
                emptycols.Add(ColumnNumber.fabric1_fpt);
            }
            if (!IsEmptyCell(row, ColumnNumber.fabric2_quality) || 
                !IsEmptyCell(row, ColumnNumber.fabric2_impression) || 
                !IsEmptyCell(row, ColumnNumber.fabric2_baseColor) || 
                !IsEmptyCell(row, ColumnNumber.fabric2_printCode) ||
                !IsEmptyCell(row, ColumnNumber.fabric2_fpt)
                ) {
                if (IsEmptyCell(row, ColumnNumber.fabric2_quality)) {
                    emptycols.Add(ColumnNumber.fabric2_quality);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric2_impression)) {
                    emptycols.Add(ColumnNumber.fabric2_impression);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric2_baseColor)) {
                    emptycols.Add(ColumnNumber.fabric2_baseColor);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric2_printCode)) {
                    emptycols.Add(ColumnNumber.fabric2_printCode);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric2_fpt)) {
                    emptycols.Add(ColumnNumber.fabric2_fpt);
                }
            }
            if (!IsEmptyCell(row, ColumnNumber.fabric3_quality) || 
                !IsEmptyCell(row, ColumnNumber.fabric3_impression) || 
                !IsEmptyCell(row, ColumnNumber.fabric3_baseColor) || 
                !IsEmptyCell(row, ColumnNumber.fabric3_printCode) ||
                !IsEmptyCell(row, ColumnNumber.fabric3_fpt)) {
                if (IsEmptyCell(row, ColumnNumber.fabric3_quality)) {
                    emptycols.Add(ColumnNumber.fabric3_quality);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric3_impression)) {
                    emptycols.Add(ColumnNumber.fabric3_impression);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric3_baseColor)) {
                    emptycols.Add(ColumnNumber.fabric3_baseColor);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric3_printCode)) {
                    emptycols.Add(ColumnNumber.fabric3_printCode);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric3_fpt)) {
                    emptycols.Add(ColumnNumber.fabric3_fpt);
                }
            }
            if (!IsEmptyCell(row, ColumnNumber.fabric4_quality) || 
                !IsEmptyCell(row, ColumnNumber.fabric4_impression) || 
                !IsEmptyCell(row, ColumnNumber.fabric4_baseColor) || 
                !IsEmptyCell(row, ColumnNumber.fabric4_printCode) ||
                !IsEmptyCell(row, ColumnNumber.fabric4_fpt)) {
                if (IsEmptyCell(row, ColumnNumber.fabric4_quality)) {
                    emptycols.Add(ColumnNumber.fabric4_quality);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric4_impression)) {
                    emptycols.Add(ColumnNumber.fabric4_impression);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric4_baseColor)) {
                    emptycols.Add(ColumnNumber.fabric4_baseColor);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric4_printCode)) {
                    emptycols.Add(ColumnNumber.fabric4_printCode);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric4_fpt)) {
                    emptycols.Add(ColumnNumber.fabric4_fpt);
                }            
            }
            if (!IsEmptyCell(row, ColumnNumber.fabric5_quality) || 
                !IsEmptyCell(row, ColumnNumber.fabric5_impression) || 
                !IsEmptyCell(row, ColumnNumber.fabric5_baseColor) || 
                !IsEmptyCell(row, ColumnNumber.fabric5_printCode) ||
                !IsEmptyCell(row, ColumnNumber.fabric5_fpt)) {
                if (IsEmptyCell(row, ColumnNumber.fabric5_quality)) {
                    emptycols.Add(ColumnNumber.fabric5_quality);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric5_impression)) {
                    emptycols.Add(ColumnNumber.fabric5_impression);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric5_baseColor)) {
                    emptycols.Add(ColumnNumber.fabric5_baseColor);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric5_printCode)) {
                    emptycols.Add(ColumnNumber.fabric5_printCode);
                }
                if (IsEmptyCell(row, ColumnNumber.fabric5_fpt)) {
                    emptycols.Add(ColumnNumber.fabric5_fpt);
                }
            }
            if (IsEmptyCell(row, ColumnNumber.dropName)) {
                emptycols.Add(ColumnNumber.dropName);
            }
            if (IsEmptyCell(row, ColumnNumber.mrpRange)) {
                emptycols.Add(ColumnNumber.mrpRange);
            }
            //if (IsEmptyCell(row, ColumnNumber.bmTarget)) {
            //    emptycols.Add(ColumnNumber.bmTarget);
            //}
            if (IsEmptyCell(row, ColumnNumber.sizeType)) {
                emptycols.Add(ColumnNumber.sizeType);
            }
            if (IsEmptyCell(row, ColumnNumber.bodyCode)) {
                emptycols.Add(ColumnNumber.bodyCode);
            }
            if (IsEmptyCell(row, ColumnNumber.dataSource)) {
                emptycols.Add(ColumnNumber.dataSource);
            }
            if (IsEmptyCell(row, ColumnNumber.dataSourceDetails)) {
                emptycols.Add(ColumnNumber.dataSourceDetails);
            }
            if (IsEmptyCell(row, ColumnNumber.color)) {
                emptycols.Add(ColumnNumber.color);
            }
            if (IsEmptyCell(row, ColumnNumber.isWashReferenced)) {
                emptycols.Add(ColumnNumber.isWashReferenced);
            }
            if (IsEmptyCell(row, ColumnNumber.pdpCatalogCallouts)) {
                emptycols.Add(ColumnNumber.pdpCatalogCallouts);
            }
            if (IsEmptyCell(row, ColumnNumber.source)) {
                emptycols.Add(ColumnNumber.source);
            }

            notifier.NotifyForEmptyCells(row, emptycols);
            return (emptycols.Count == 0) ? false : true;
        }

        public void ValidateHandovers(List<Handover> handoverlist)
        {
            List<ValidatorResult> result = messenger.GetValidationInfo(handoverlist);            
            int row = 2;
            foreach(ValidatorResult vr in result)
            {
                System.Diagnostics.Debug.WriteLine(JsonConvert.SerializeObject(vr.isValid, Formatting.Indented));
                if (!vr.allok)
                {
                    if(vr.isValid.ContainsKey("sizetype") == true && !vr.isValid["sizetype"])
                    {
                        decorator.HighlightErrorAtCell(row, ColumnNumber.sizeType, "Rejected value of Sizetype");
                    }
                    if (vr.isValid.ContainsKey("bag") == true && !vr.isValid["bag"])
                    {
                        decorator.HighlightErrorAtCell(row, ColumnNumber.brand, "Rejected Value by Validation Master BAG; " +
                            "Please correct Brand, Article Type and Gender Combination.");
                        decorator.HighlightErrorAtCell(row, ColumnNumber.articleType, "Rejected Value by Validation Master BAG; " +
                            "Please correct Brand, Article Type and Gender Combination.");
                        decorator.HighlightErrorAtCell(row, ColumnNumber.gender, "Rejected Value by Validation Master BAG; " +
                            "Please correct Brand, Article Type and Gender Combination.");
                    }
                    if (vr.isValid.ContainsKey("quantity") == true && !vr.isValid["quantity"])
                    {
                        decorator.HighlightErrorAtCell(row, ColumnNumber.quantity, "Rejected Value by Validation Master MoQ; " +
                            "Please correct Brand, Article Type, Gender, Subcategory and Quantity Combination.");
                    }
                    if (vr.isValid.ContainsKey("cluster") == true && !vr.isValid["cluster"])
                    {
                        decorator.HighlightErrorAtCell(row, ColumnNumber.cluster, "Rejected Value by Validation Master Cluster; " +
                            "Please correct Gender and Cluster Combination.");
                    }
                    if (vr.isValid.ContainsKey("subcategory") == true && !vr.isValid["subcategory"])
                    {
                        decorator.HighlightErrorAtCell(row, ColumnNumber.subcategory, "Rejected Value by Validation Master Subcategory; " +
                            "Please correct Brand, Article Type, Gender and Subcategory Combination.");
                    }
                    if (vr.isValid.ContainsKey("bmtarget") == true && !vr.isValid["bmtarget"])
                    {
                        decorator.HighlightErrorAtCell(row, ColumnNumber.bmTarget, "Rejected Value by Validation Master BmTarget; " +
                            "Please correct Brand, Article Type, Gender, Repeated and BM Target Combination.");
                    }
                    if (vr.isValid.ContainsKey("bodycode") == true && !vr.isValid["bodycode"])
                    {
                        decorator.HighlightErrorAtCell(row, ColumnNumber.bodyCode, "Rejected Value by Validation Master BodyCode; " +
                            "Please correct Article Type, Gender and BodyCode Combination.");
                    }
                    if (vr.isValid.ContainsKey("color") == true && !vr.isValid["color"])
                    {
                        decorator.HighlightErrorAtCell(row, ColumnNumber.color, "Rejected Value by Validation Master Color; " +
                            "Please correct Gender and Color Combination.");
                    }

                } else
                {
                    decorator.ClearAllErrors(row);
                }
                ++row;
            }
        }
    }
}
