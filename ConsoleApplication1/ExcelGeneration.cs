﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using ConsoleApplication1;
using ExportValidationConsole;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExportValidation.Common
{
    public static class ExcelGeneration
    {
        private static Excel.Application CreateExcelObj()
        {
            object obj;
            obj = null;
            try
            {
                //Создаём приложение.
                Excel.Application objExcel = new Excel.Application();
                objExcel.Workbooks.Add();
                obj = objExcel;

            }
            catch (Exception ex)
            {
                throw new Exception("Ошибка создания экземпляра MS Excel");
            }
            return (obj as Excel.Application);
        }

        private static void FormatDescription(Excel.Range range)
        {
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.RowHeight = 18;
            range.Font.Size = 12;
            range.Font.Bold = true;
            range.EntireColumn.AutoFit();

        }

  

        private static void FormatDataArea(Excel.Range range, String data)
        {
            range.Worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, range, System.Type.Missing,
                Excel.XlYesNoGuess.xlYes, System.Type.Missing).Name = data;
            range.Select();
            range.Worksheet.ListObjects[data].TableStyle = "TableStyleMedium2";
            range.EntireColumn.AutoFit();
        }

        private static void FormatSheet(Excel.Worksheet sheet, QueryData obj)
        {
            //sheet.PageSetup.PrintGridlines = false;
            // sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            // sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
            // sheet.PageSetup.RightFooter = "Дата: &DD Стр &PP из &NN";
            // sheet.PageSetup.RightHeader = obj.ProjectName + " - " + obj.ValidationRule;
            // sheet.PageSetup.Zoom = false;
            // sheet.PageSetup.LeftHeader = "e-CRF";
            // sheet.PageSetup.TopMargin = 50;
            // sheet.PageSetup.BottomMargin = 50;
            // sheet.PageSetup.HeaderMargin = 20;
            // sheet.PageSetup.FooterMargin = 20;
            // sheet.PageSetup.RightMargin = 10;
            // sheet.PageSetup.LeftMargin = 50;
            // sheet.PageSetup.Order = Excel.XlOrder.xlOverThenDown;
            sheet.Columns.EntireColumn.AutoFit();
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[obj.Data.Rows.Count, 3]].NumberFormat = "@";
        }

        private static List<QueryData> SortData(List<QueryData> data)
        {
            QueryData tmp = null;
            try
            {


                for (int i = 0; i < data.Count - 1; i++)
                {
                    for (int j = 0; j < data.Count - 1; j++)
                    {
                        if (Convert.ToInt32(data[j].NameList) < Convert.ToInt32(data[j + 1].NameList))
                        {
                            tmp = data[j];
                            data[j] = data[j + 1];
                            data[j + 1] = tmp;

                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {

            }
            return data;
        }

        public static string GenerateDocument(string filePath,ReturnProc res)
        {
            Excel.Application ExcelApp;
            Excel.Worksheet ExcelSheet;
            Excel.Workbook ExcelWorkbook;
            Excel.Workbooks ExcelWorkbooks;
            Excel.Range ExcelRange;
            int rowsCount;
            int colsCount;
            Log.Write("Start GenerateDocument");
            try
            {
                if (String.IsNullOrEmpty(filePath))
                {
                    filePath = Path.Combine(filePath  + res.Data[0].ProjectName + "_Validation_" + DateTime.Now.ToShortDateString());
                    //if (Environment.OSVersion.Version.Major >= 6)
                    //{
                    //    filePath = Directory.GetParent(filePath).FullName;
                    //}
                }
                var filename = res.Data[0].ProjectName + "_Validation_" + DateTime.Now.ToShortDateString() + ".xlsx";

      
                ExcelApp = CreateExcelObj();
                ExcelWorkbooks = ExcelApp.Workbooks;
                ExcelApp.ScreenUpdating = false;
                ExcelApp.DisplayAlerts = false;
                ExcelWorkbook = ExcelWorkbooks.Add();

                res.Data = SortData(res.Data);

                foreach (var itemData in res.Data)
                {
                    Log.Write(itemData.NameList + " " + itemData.ValidationRule);
                    ExcelSheet = ExcelWorkbook.Sheets.Add();
                    ExcelSheet.Name = itemData.NameList;
                    FormatSheet(ExcelSheet, itemData);
                    rowsCount = itemData.Data.Rows.Count;
                    colsCount = itemData.Data.Columns.Count;

                    ExcelSheet.Cells[1, 1] = "Проект:";
                    ExcelSheet.Cells[1, 2] = itemData.ProjectName;
                    ExcelSheet.Range[ExcelSheet.Cells[1, 2], ExcelSheet.Cells[1, colsCount]].Merge();
                    ExcelSheet.Cells[2, 1] = "Правило:";
                    ExcelSheet.Cells[2, 2] = itemData.ValidationRule;
                    ExcelSheet.Range[ExcelSheet.Cells[2, 2], ExcelSheet.Cells[2, colsCount]].Merge();
                    ExcelSheet.Cells[3, 1] = "Описание:";
                    ExcelSheet.Cells[3, 2] = itemData.Description;
                    ExcelSheet.Range[ExcelSheet.Cells[3, 2], ExcelSheet.Cells[3, colsCount]].Merge();

                    object[,] dataSet = new object[itemData.Data.Rows.Count, itemData.Data.Columns.Count];




                    for (int i = 1; i <= itemData.FieldsName.Count; i++)
                    {
                        ExcelSheet.Cells[5, i] = itemData.FieldsName[i - 1];
                    }

                    for (int i = 0; i < itemData.Data.Rows.Count; i++)
                    {
                        for (int j = 0; j < itemData.FieldsName.Count; j++)
                        {
                            dataSet[i, j] = itemData.Data.Rows[i][j].ToString();
                        }
                    }
                    Excel.Range rng =
                        ExcelSheet.Range[
                            ExcelSheet.Cells[6, 1],
                            ExcelSheet.Cells[5 + itemData.Data.Rows.Count, itemData.FieldsName.Count]];
                    rng.Value = dataSet;

                    FormatDataArea(ExcelSheet.Range[ExcelSheet.Cells[5, 1], ExcelSheet.Cells[5 + rowsCount, colsCount]], itemData.ValidationRule);
                    FormatDescription(ExcelSheet.Range[ExcelSheet.Cells[1, 1], ExcelSheet.Cells[3, 2]]);
                    ExcelSheet.Range[ExcelSheet.Cells[3, 2], ExcelSheet.Cells[3, colsCount]].WrapText = true;
                    ExcelSheet.Range[ExcelSheet.Cells[3, 2], ExcelSheet.Cells[3, colsCount]].RowHeight = 60;

                }

                //Adding Index
                object[,] dataIndex = new object[res.Index.Count, 4];
                for (int j = 0; j < res.Index.Count; j++)
                {
                    dataIndex[j, 0] = res.Index[j].NameList;
                    dataIndex[j, 1] = res.Index[j].ValidationRule;
                    dataIndex[j, 2] = res.Index[j].Description;
                    dataIndex[j, 3] = res.Index[j].SelectCommand;
                }
                ExcelSheet = ExcelWorkbook.Sheets.Add();
                ExcelSheet.Name = "Index";
                FormatSheet(ExcelSheet, res.Data[0]);

                ExcelSheet.Cells[1, 1] = "Проект:";
                ExcelSheet.Cells[1, 2] = res.Data[0].ProjectName;
                ExcelSheet.Cells[2, 1] = "Описание:";
                ExcelSheet.Range[ExcelSheet.Cells[2, 1], ExcelSheet.Cells[2, 3]].Merge();
                FormatDescription(ExcelSheet.Range[ExcelSheet.Cells[1, 1], ExcelSheet.Cells[2, 2]]);
                ExcelSheet.Cells[4, 1] = "Имя листа";
                ExcelSheet.Cells[4, 2] = "Правило валидации";
                ExcelSheet.Cells[4, 3] = "Описание правила для поиска ошибок";
               ExcelSheet.Cells[4, 4] = "SQL";

                Excel.Range rngIndex = ExcelSheet.Range[ExcelSheet.Cells[5, 1], ExcelSheet.Cells[4 + res.Index.Count, 4]];
                rngIndex.Value = dataIndex;
                FormatDataArea(ExcelSheet.Range[ExcelSheet.Cells[4, 1], ExcelSheet.Cells[4 + res.Index.Count, 4]], "index");
                ExcelSheet.Range[ExcelSheet.Cells[5, 3], ExcelSheet.Cells[4 + res.Index.Count, 3]].ColumnWidth = 90;
                ExcelSheet.Range[ExcelSheet.Cells[5, 3], ExcelSheet.Cells[4 + res.Index.Count, 3]].WrapText = true;

                //End Creating Index

                ExcelWorkbook.SaveAs();
                ExcelWorkbook.SaveAs(filePath  + filename, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);


                while (Marshal.ReleaseComObject(ExcelWorkbook) > 0)
                { }
                while (Marshal.ReleaseComObject(ExcelWorkbooks) > 0)
                { }


                ExcelApp.Quit();

                while (Marshal.ReleaseComObject(ExcelApp) > 0)
                { }
                Log.Write("Document created successfully !");
                return filePath + "\\" + filename;


            }
            catch (Exception ex)
            {
                Log.Write(ex.Data + "\r\n" + ex.Message + "\r\n" + ex.Source + "\r\n" + ex.InnerException + "\r\n" +
                                ex.StackTrace);
                return null;
            }
            finally
            {

                GC.Collect();
                Log.Write("Exit GenerateDocument");
            }
        }

        public static string GenerateDocument2(string filePath, ReturnProc res)
        {
            Excel.Application ExcelApp;
            Excel.Worksheet ExcelSheet;
            Excel.Workbook ExcelWorkbook;
            Excel.Workbooks ExcelWorkbooks;
            Excel.Range ExcelRange;
            int rowsCount;
            int colsCount;
            Log.Write("Start GenerateDocument2");
            try
            {
                if (String.IsNullOrEmpty(filePath))
                {
                    filePath =
                        Path.Combine(filePath  + res.Data[0].ProjectName + "_Export_" + DateTime.Now.ToShortDateString());
                    //if (Environment.OSVersion.Version.Major >= 6)
                    //{
                    //    filePath = Directory.GetParent(filePath).FullName;
                    //}
                }
                var filename = res.Data[0].ProjectName + "_Export_" + DateTime.Now.ToShortDateString() + ".xlsx";

                var indexDocument = res.Index;

                ExcelApp = CreateExcelObj();
                ExcelWorkbooks = ExcelApp.Workbooks;
                ExcelApp.ScreenUpdating = false;
                ExcelApp.DisplayAlerts = false;
                ExcelWorkbook = ExcelWorkbooks.Add();

                res.Data = SortData(res.Data);

                foreach (var itemData in res.Data)
                {
                    Log.Write(itemData.NameList + " " + itemData.ValidationRule);
                    ExcelSheet = ExcelWorkbook.Sheets.Add();
                    ExcelSheet.Name = itemData.NameList;
                    FormatSheet(ExcelSheet, itemData);
                    rowsCount = itemData.Data.Rows.Count;
                    colsCount = itemData.Data.Columns.Count;

                    //ExcelSheet.Cells[1, 1] = "Проект:";
                    //ExcelSheet.Cells[1, 2] = itemData.ProjectName;
                    //ExcelSheet.Range[ExcelSheet.Cells[1, 2], ExcelSheet.Cells[1, colsCount]].Merge();
                    //ExcelSheet.Cells[2, 1] = "Правило:";
                    //ExcelSheet.Cells[2, 2] = itemData.ValidationRule;
                    //ExcelSheet.Range[ExcelSheet.Cells[2, 2], ExcelSheet.Cells[2, colsCount]].Merge();
                    //ExcelSheet.Cells[3, 1] = "Описание:";
                    //ExcelSheet.Cells[3, 2] = itemData.Description;
                    //ExcelSheet.Range[ExcelSheet.Cells[3, 2], ExcelSheet.Cells[3, colsCount]].Merge();

                    object[,] dataSet = new object[itemData.Data.Rows.Count, itemData.Data.Columns.Count];




                    for (int i = 1; i <= itemData.FieldsName.Count; i++)
                    {
                        ExcelSheet.Cells[1, i] = itemData.FieldsName[i - 1];
                    }

                    for (int i = 0; i < itemData.Data.Rows.Count; i++)
                    {
                        for (int j = 0; j < itemData.FieldsName.Count; j++)
                        {
                            dataSet[i, j] = itemData.Data.Rows[i][j].ToString();
                        }
                    }
                    Excel.Range rng =
                        ExcelSheet.Range[
                            ExcelSheet.Cells[2, 1],
                            ExcelSheet.Cells[1 + itemData.Data.Rows.Count, itemData.FieldsName.Count]];
                    rng.Value = dataSet;

                    FormatDataArea(
                        ExcelSheet.Range[ExcelSheet.Cells[1, 1], ExcelSheet.Cells[1 + rowsCount, colsCount]],
                        itemData.ValidationRule);
                    //FormatDescription(ExcelSheet.Range[ExcelSheet.Cells[1, 1], ExcelSheet.Cells[3, 2]]);
                    // ExcelSheet.Range[ExcelSheet.Cells[3, 2], ExcelSheet.Cells[3, colsCount]].WrapText = true;
                    // ExcelSheet.Range[ExcelSheet.Cells[3, 2], ExcelSheet.Cells[3, colsCount]].RowHeight = 60;

                }

                //Adding Index
                object[,] dataIndex = new object[indexDocument.Count, 4];
                for (int j = 0; j < indexDocument.Count; j++)
                {
                    dataIndex[j, 0] = indexDocument[j].NameList;
                    dataIndex[j, 1] = indexDocument[j].ValidationRule;
                    dataIndex[j, 2] = indexDocument[j].Description;
                    dataIndex[j, 3] = indexDocument[j].SelectCommand;
                }
                ExcelSheet = ExcelWorkbook.Sheets.Add();
                ExcelSheet.Name = "Index";
                FormatSheet(ExcelSheet, res.Data[0]);

                ExcelSheet.Cells[1, 1] = "Проект:";
                ExcelSheet.Cells[1, 2] = res.Data[0].ProjectName;
                ExcelSheet.Cells[2, 1] = "Описание:";
                ExcelSheet.Range[ExcelSheet.Cells[2, 1], ExcelSheet.Cells[2, 3]].Merge();
                FormatDescription(ExcelSheet.Range[ExcelSheet.Cells[1, 1], ExcelSheet.Cells[2, 2]]);
                ExcelSheet.Cells[4, 1] = "Имя листа";
                ExcelSheet.Cells[4, 2] = "Правило валидации";
                ExcelSheet.Cells[4, 3] = "Описание правила для поиска ошибок";
                ExcelSheet.Cells[4, 4] = "SQL";

                Excel.Range rngIndex =
                    ExcelSheet.Range[ExcelSheet.Cells[5, 1], ExcelSheet.Cells[4 + indexDocument.Count, 4]];
                rngIndex.Value = dataIndex;
                FormatDataArea(ExcelSheet.Range[ExcelSheet.Cells[4, 1], ExcelSheet.Cells[4 + indexDocument.Count, 4]],
                    "index");
                ExcelSheet.Range[ExcelSheet.Cells[5, 3], ExcelSheet.Cells[4 + indexDocument.Count, 4]].ColumnWidth = 90;
                ExcelSheet.Range[ExcelSheet.Cells[5, 3], ExcelSheet.Cells[4 + indexDocument.Count, 4]].WrapText = true;

                //End Creating Index

                ExcelWorkbook.SaveAs();
                ExcelWorkbook.SaveAs(filePath + "\\" + filename, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing,
                    Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                    Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);


                while (Marshal.ReleaseComObject(ExcelWorkbook) > 0)
                {
                }
                while (Marshal.ReleaseComObject(ExcelWorkbooks) > 0)
                {
                }


                ExcelApp.Quit();

                while (Marshal.ReleaseComObject(ExcelApp) > 0)
                {
                }
                Log.Write("Document created successfully !");
                return filePath  + filename;
            }
            catch (Exception ex)
            {
                Log.Write(ex);
                return null;
            }
            finally
            {
                
                GC.Collect();
                Log.Write("Exit GenerateDocument2");
            }
        }
    }
}
