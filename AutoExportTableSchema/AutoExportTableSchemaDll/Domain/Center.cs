﻿using AutoExportTableSchema.Model;
using AutoExportTableSchema.SqlConnect;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;

namespace AutoExportTableSchema.Domain
{
    public class Center
    {
        public static readonly string Downloads = new Syroot.Windows.IO.KnownFolder(Syroot.Windows.IO.KnownFolderType.Downloads).Path;
        ExcelPackage excel { get; set; }
        ExcelWorksheet workSheet { get; set; }
        public void Run(string ConnectString, string FileName,bool bChkDes)
        {
            excel = new ExcelPackage();
            setExcelTitle();
            setExcelContext(new SqlData(ConnectString), bChkDes);
            runFinish(FileName);
        }
        public void setMergeContext(ISqlData sqlData)
        {
            int i = 0;
            int j = 0;
            var tempList = sqlData.lstAttributes().Select(c => c.TableName).Distinct().ToArray();

        }
        public void setExcelTitle()
        {
            workSheet = excel.Workbook.Worksheets.Add("總表");
            changeBackgroundColor(workSheet, $"A1:I1", Color.Yellow);
            workSheet.Cells[1, 1].Value = "SchemaName";
            workSheet.Cells[1, 2].Value = "TableName";
            workSheet.Cells[1, 3].Value = "Column Name";
            workSheet.Cells[1, 4].Value = "Column Type";
            workSheet.Cells[1, 5].Value = "Max Length";
            workSheet.Cells[1, 6].Value = "IsNull";
            workSheet.Cells[1, 7].Value = "IsPrimaryKey";
            workSheet.Cells[1, 8].Value = "describe";
            workSheet.Cells[1, 9].Value = "describe2";
            workSheet.Cells["A1:I1"].AutoFilter = true;
            workSheet.Cells[$"A2:I2"].Merge = true;
            workSheet.Cells["A2:I2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells["A2:I2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            workSheet.Cells[$"A2"].Style.Font.Size = 20;
            workSheet.Cells[$"A2"].Style.Font.Bold = true;
            changeBackgroundColor(workSheet, "A2:I2", Color.FromArgb(255, 127, 80));
            workSheet.Row(2).Height = 25;

          

        }
        public static bool SqlConnectionTest(string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    return true;
                }
                catch (SqlException)
                {
                    return false;
                }
            }
        }
        public static void CreatTempelteExcel()
        {
            string FileName = "Templete";
            ExcelPackage excel = new ExcelPackage();
            ExcelWorksheet workSheet = excel.Workbook.Worksheets.Add("Sheet1");
            workSheet.Cells[1, 1].Value = "伺服器名稱";
            workSheet.Cells[1, 2].Value = "資料庫名稱";
            workSheet.Cells[1, 3].Value = "帳號";
            workSheet.Cells[1, 4].Value = "密碼";

            workSheet.Cells[$"A1:D1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[$"A1:D1"].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            workSheet.Cells["A1:D1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            var modelTable = workSheet.Cells["A1:D1"];
            // Assign borders
            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            workSheet.Cells["A1:D1"].AutoFitColumns();
            string strPath = Path.Combine(Downloads, $"{FileName}_{DateTime.Now.ToString("yymmddhhmmss")}.xlsx");
            FileStream objFileStrm = File.Create(strPath);
            objFileStrm.Close();
            File.WriteAllBytes(strPath, excel.GetAsByteArray());
            excel.Dispose();
            Process.Start(strPath);
        }
        public IEnumerable<Tempelte> ReadTempleteExecl(string FilePath)
        {
            List<Tempelte> lst = new List<Tempelte>();
            Excel excel = new Excel(FilePath, "Sheet1");
            for (int i = 0; i < excel.sheet.Dimension.End.Row; i++)
            {
                var E1 = excel.sheet.Cells[i + 2, 1].Value;
                var E2 = excel.sheet.Cells[i + 2, 2].Value;
                var E3 = excel.sheet.Cells[i + 2, 3].Value;
                var E4 = excel.sheet.Cells[i + 2, 4].Value;
                if (E1 != null && E2 != null && E3 != null && E4 != null)
                {
                    lst.Add(new Tempelte(E1, E2, E3, E4));
                }
            }
            return lst;
        }
        public void setExcelContext(ISqlData sqlData,bool bExistDes)
        {
            int i = 0;
            var tempList = sqlData.lstAttributes().Select(c => c.TableName).Distinct().ToList();
            //workSheet.Cells[$"B{min}:Z{max}"].Merge = true;

            //Excel總行數 =  sqlData.lstAttributes().Count() + Excel標題  +  sqlData.lstAttributes().Select(c => c.TableName).Distinct().Count();
            foreach (var item in tempList)
            {
                workSheet.Cells[$"A{2 + i}:I{2 + i}"].Merge = true;
                workSheet.Cells[2 + i, 1].Value = item;
                workSheet.Cells[2 + i, 2].Value = "";
                workSheet.Cells[2 + i, 3].Value = "";
                workSheet.Cells[2 + i, 4].Value = "";
                workSheet.Cells[2 + i, 5].Value = "";
                workSheet.Cells[2 + i, 6].Value = "";
                workSheet.Cells[2 + i, 7].Value = "";

                workSheet.Cells[$"A{2 + i}"].Value = workSheet.Cells[$"B{2 + i}"].Value;
                workSheet.Cells[$"A{2 + i}"].Style.Font.Size = 20;
                workSheet.Cells[$"A{2 + i}"].Style.Font.Bold = true;
                workSheet.Cells[$"A{2 + i}:I{2 + i}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[$"A{2 + i}:I{2 + i}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                changeBackgroundColor(workSheet, $"A{2 + i}:I{2 + i}", Color.FromArgb(255, 127, 80));
                workSheet.Row(3 + i).Height = 25;
                var tableName = sqlData.lstAttributes().Where(c => c.TableName == item).ToList();
                foreach (var item2 in tableName)
                {
                    i++;
                    workSheet.Cells[2 + i, 1].Value = item2.SchemaName;
                    workSheet.Cells[2 + i, 2].Value = item2.TableName;
                    workSheet.Cells[2 + i, 3].Value = item2.ColumnName;
                    workSheet.Cells[2 + i, 4].Value = item2.ColumnType;
                    workSheet.Cells[2 + i, 5].Value = item2.MaxLength;
                    workSheet.Cells[2 + i, 6].Value = item2.IsNull;
                    workSheet.Cells[2 + i, 7].Value = item2.IsPrimaryKey;
                    if (bExistDes == true)
                    {
                        workSheet.Cells[2 + i, 8].Value = item2.ColumnDescription;
                    }
                    changeBackgroundColor(workSheet, $"A{2 + i}:I{2 + i}", Color.FromArgb(255, 228, 225));
                }
                i++;
            }



            var modelTable = workSheet.Cells[$"A1:I{sqlData.lstAttributes().Count() + tempList.Count() + 1}"];
            workSheet.Cells[$"A1:I{sqlData.lstAttributes().Count()}"].AutoFitColumns();
            workSheet.Column(8).Width = 30;
            // Assign borders
            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            IExcelSkill skill = new ExcelSkill(workSheet);
            skill.setMergeValue(tempList);
            var tempeee = skill.getMergeRange(out int last);
            foreach (var item in tempeee)
            {
                workSheet.Cells[$"I{item.Min}:I{item.Max}"].Merge = true;
                workSheet.Cells[$"A{item.Min}:I{item.Max}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Cells[$"A{item.Min}:I{item.Max}"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                workSheet.Cells[$"I{item.Min}:I{item.Max}"].Style.Font.Name = "微軟正黑體";
            }
            workSheet.Cells[$"I{last + 1}:I{i + 1}"].Merge = true;
            workSheet.Column(9).Width = 30;

        }
        public void changeBackgroundColor(ExcelWorksheet workSheet, string Range, Color color)
        {

            workSheet.Cells[Range].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[Range].Style.Fill.BackgroundColor.SetColor(color);
        }
        public void runFinish(string FileName)
        {
            string p_strPath = Path.Combine(Downloads, $"{FileName}_{DateTime.Now.ToString("yymmddhhmmss")}.xlsx");
            FileStream objFileStrm = File.Create(p_strPath);
            objFileStrm.Close();

            // Write content to excel file 
            File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
            //Close Excel package
            excel.Dispose();
            Process.Start(p_strPath);
        }
        public static void openFile(string strFIle)
        {
            Process.Start(strFIle);
        }

    }
}
