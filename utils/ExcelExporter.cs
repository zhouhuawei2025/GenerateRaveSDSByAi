using GenerateRaveSDSByAi.dataModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace GenerateRaveSDSByAi.utils
{
    internal class ExcelExporter
    {
        public static void ExportFormToExcel(List<Form> formList, string outputPath)
        {
            using (var package = new ExcelPackage())
            {
                // ========== Sheet1：Module表 ==========
                var sheet1 = package.Workbook.Worksheets.Add("Forms");
                // 设置表头
                string[] sheet1Headers = { 
                    "OID", "Ordinal", "DraftFormName", "DraftFormActive", "HelpText", 
                    "IsTemplate", "IsSignatureRequired", "IsEproForm", "ViewRestrictions",   "EntryRestrictions",
                    "LogDirection",    "DDEOption",   "ConfirmationStyle",   "LinkFolderOID", "LinkFormOID"};
                for (int col = 0; col < sheet1Headers.Length; col++)
                {
                    sheet1.Cells[1, col + 1].Value = sheet1Headers[col];
                    sheet1.Cells[1, col + 1].Style.Font.Bold = true; // 表头加粗
                }

                // 填充数据
                int sheet1Row = 2;
                int num = 1;
                foreach (var form in formList)
                {
                    sheet1.Cells[sheet1Row, 1].Value = form.OID;
                    sheet1.Cells[sheet1Row, 2].Value = num.ToString();
                    sheet1.Cells[sheet1Row, 3].Value = form.Name;
                    sheet1.Cells[sheet1Row, 4].Value = "TRUE";
                    sheet1.Cells[sheet1Row, 6].Value = "FALSE";
                    sheet1.Cells[sheet1Row, 7].Value = "TRUE";
                    sheet1.Cells[sheet1Row, 8].Value = "FALSE";
                    sheet1.Cells[sheet1Row, 11].Value = form.LogDirection;
                    sheet1.Cells[sheet1Row, 12].Value = "MustNotDDE";
                    sheet1.Cells[sheet1Row, 13].Value = "NoLink";
                    num++;
                    sheet1Row++;
                }
                sheet1.Cells.AutoFitColumns(); // 自动调整列宽

                // 保存Excel文件
                FileInfo file = new FileInfo(outputPath);
                package.SaveAs(file);
                Debug.WriteLine($"Excel导出成功！路径：{outputPath}");
            }
        }
        public static void ExportDataDictionaryToExcel(List<DataDictionary> dictList, string outputPath)
        {
            using (var package = new ExcelPackage())
            {
                // ========== Sheet1：DataDictionary主表 ==========
                var sheet1 = package.Workbook.Worksheets.Add("DataDictionaries");
                // 设置Sheet1表头
                string[] sheet1Headers = { "DataDictionaryName", "OID"};
                for (int col = 0; col < sheet1Headers.Length; col++)
                {
                    sheet1.Cells[1, col + 1].Value = sheet1Headers[col];
                    sheet1.Cells[1, col + 1].Style.Font.Bold = true; // 表头加粗                   
                }
                
                int sheet1Row = 2;
                foreach (var dict in dictList)
                {
                    sheet1.Cells[sheet1Row, 1].Value = dict.Name;
                    sheet1.Cells[sheet1Row, 2].Value = dict.OID;                    
                    sheet1Row++;
                }
                sheet1.Cells.AutoFitColumns(); // 自动调整列宽


                // ========== Sheet2：DataDictionaryEntry明细表 ==========
                var sheet2 = package.Workbook.Worksheets.Add("DataDictionaryEntries");
                // 设置Sheet2表头
                string[] sheet2Headers = {"DataDictionaryName",    "CodedData",   "Ordinal", "UserDataString",  "Specify"};
                for (int col = 0; col < sheet2Headers.Length; col++)
                {
                    sheet2.Cells[1, col + 1].Value = sheet2Headers[col];
                    sheet2.Cells[1, col + 1].Style.Font.Bold = true;
                }

                // 填充Sheet2数据               

                int sheet2Row = 2;
                foreach (var dict in dictList)
                {
                    int ordinal = 1;
                    foreach (var entry in dict.DataDictionaryEntries)
                    {
                        sheet2.Cells[sheet2Row, 1].Value = dict.Name;
                        sheet2.Cells[sheet2Row, 2].Value = entry.EntryOID.ToString();
                        sheet2.Cells[sheet2Row, 3].Value = ordinal;
                        sheet2.Cells[sheet2Row, 4].Value = entry.ItemDataString;
                        sheet2.Cells[sheet2Row, 5].Value = entry.IsSpecify ? "TRUE" : "FALSE";                       
                        sheet2Row++;
                        ordinal++;
                    }
                }
                sheet2.Cells.AutoFitColumns(); // 自动调整列宽

                // 保存Excel文件
                FileInfo file = new FileInfo(outputPath);
                package.SaveAs(file);
                Debug.WriteLine($"Excel导出成功！路径：{outputPath}");
            }
        }



        //public static void ExportAnalyteToExcel(List<List<Field>> fieldList, string outputPath)
        //{
        //    using (var package = new ExcelPackage())
        //    {
        //        var sheet1 = package.Workbook.Worksheets.Add("LabKey");
        //        // 设置表头
        //        string[] sheet1Headers = new string[] { "LabKey", "KeyDescription", "StandardUnit" };

        //        for (int col = 0; col < sheet1Headers.Length; col++)
        //        {
        //            sheet1.Cells[1, col + 1].Value = sheet1Headers[col];
        //            sheet1.Cells[1, col + 1].Style.Font.Bold = true; // 表头加粗
        //        }

        //        // 填充数据
        //        int sheet1Row = 2;
        //        var fields = from u in fieldList
        //                     from v in u.Where(g => g.IsLab == true)
        //                     select v;

        //        foreach (var field in fields)
        //        {
        //            sheet1.Cells[sheet1Row, 1].Value = field.FieldOID;
        //            sheet1.Cells[sheet1Row, 2].Value = field.FieldName;
        //            sheet1Row++;
        //        }
        //        sheet1.Cells.AutoFitColumns(); // 自动调整列宽

        //        // 保存Excel文件
        //        FileInfo file = new FileInfo(outputPath);
        //        package.SaveAs(file);
        //        Debug.WriteLine($"Excel导出成功！路径：{outputPath}");
        //    }
        //}
        public static void ExportFieldToExcel(List<List<Field>> fieldList, string outputPath)
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Fields");
                // 设置表头
                string[] sheet1Headers = new string[]
                {
                    "FormOID", "FieldOID", "Ordinal", "DraftFieldNumber", "DraftFieldName",
                    "DraftFieldActive", "VariableOID", "DataFormat", "DataDictionaryName", "UnitDictionaryName",
                    "CodingDictionary", "ControlType", "AcceptableFileExtensions", "IndentLevel", "PreText",
                    "FixedUnit", "HeaderText", "HelpText", "SourceDocument", "IsLog",
                    "DefaultValue", "SASLabel", "SASFormat", "EproFormat", "IsRequired",
                    "QueryFutureDate", "IsVisible", "IsTranslationRequired", "AnalyteName", "IsClinicalSignificance",
                    "QueryNonConformance", "OtherVisits", "CanSetRecordDate", "CanSetDataPageDate", "CanSetInstanceDate",
                    "CanSetSubjectDate", "DoesNotBreakSignature", "LowerRange", "UpperRange", "NCLowerRange",
                    "NCUpperRange", "ViewRestrictions", "EntryRestrictions", "ReviewGroups", "IsVisualVerify",
                    "FDownloadedFromObjectId", "FSourceObjectId", "VDownloadedFromObjectId", "VSourceObjectId", "FSourceUrlId",
                    "VSourceUrlId", "AnalyteName_ValCol"
                };

                for (int col = 0; col < sheet1Headers.Length; col++)
                {
                    sheet1.Cells[1, col + 1].Value = sheet1Headers[col];
                    sheet1.Cells[1, col + 1].Style.Font.Bold = true; // 表头加粗
                }

                // 填充数据
                int sheet1Row = 2;
                var fields = from u in fieldList
                             from v in u
                             select v;

                foreach (var field in fields)
                {
                    sheet1.Cells[sheet1Row, 1].Value = field.FormOID;
                    sheet1.Cells[sheet1Row, 2].Value = field.FieldOID;
                    sheet1.Cells[sheet1Row, 3].Value = field.Oridinal;
                    
                    sheet1.Cells[sheet1Row, 5].Value = field.DraftFieldName;
                    sheet1.Cells[sheet1Row, 6].Value = field.DraftFieldActive ? "TRUE":"FALSE";
                    sheet1.Cells[sheet1Row, 7].Value = field.VariableOID;
                    sheet1.Cells[sheet1Row, 8].Value = field.DataFormat;
                    sheet1.Cells[sheet1Row, 9].Value = field.DataDictionaryOID;
                    
                    sheet1.Cells[sheet1Row, 11].Value = field.CodingDictionary;
                    sheet1.Cells[sheet1Row, 12].Value = field.ControlType;
                    
                    sheet1.Cells[sheet1Row, 14].Value = "0";
                    sheet1.Cells[sheet1Row, 15].Value = field.FieldName;
                    
                    sheet1.Cells[sheet1Row, 19].Value = "TRUE";
                    sheet1.Cells[sheet1Row, 20].Value = field.IsGrid ? "TRUE" : "FALSE";
                    sheet1.Cells[sheet1Row, 21].Value = field.DefaultValue;
                    sheet1.Cells[sheet1Row, 22].Value = field.SASLabel;
                    sheet1.Cells[sheet1Row, 23].Value = field.SASFormat;
                    
                    sheet1.Cells[sheet1Row, 25].Value = field.IsRequired ? "TRUE" : "FALSE";
                    sheet1.Cells[sheet1Row, 26].Value = field.QueryFutureDate ? "TRUE" : "FALSE"; ;
                    sheet1.Cells[sheet1Row, 27].Value = field.IsVisible ? "TRUE" : "FALSE";
                    sheet1.Cells[sheet1Row, 28].Value = "FALSE";
                    sheet1.Cells[sheet1Row, 29].Value = "";  //待定
                    sheet1.Cells[sheet1Row, 30].Value = field.IsLab ? "TRUE" : "FALSE";
                    sheet1.Cells[sheet1Row, 31].Value = field.QueryNonConformance ? "TRUE" : "FALSE";
                    sheet1.Cells[sheet1Row, 32].Value = "FALSE";
                    sheet1.Cells[sheet1Row, 33].Value = "FALSE";
                    sheet1.Cells[sheet1Row, 34].Value = "FALSE";
                    sheet1.Cells[sheet1Row, 35].Value = "FALSE";
                    sheet1.Cells[sheet1Row, 36].Value = "FALSE";
                    sheet1.Cells[sheet1Row, 37].Value = field.DoesNotBreakSignature ? "TRUE" : "FALSE";

                    sheet1.Cells[sheet1Row, 45].Value = "FALSE";
                   
                    sheet1Row++;
                }
                sheet1.Cells.AutoFitColumns(); // 自动调整列宽

                // 保存Excel文件
                FileInfo file = new FileInfo(outputPath);
                package.SaveAs(file);
                Debug.WriteLine($"Excel导出成功！路径：{outputPath}");
            }
        }
    }
}
