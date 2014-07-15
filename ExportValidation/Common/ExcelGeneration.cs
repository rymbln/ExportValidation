using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
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

        private static void FormatHighlitedCells(Excel.Range range)
        {
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;

            // You should specify all these parameters every time you call this method, 
            // since they can be overridden in the user interface. 
            currentFind = range.Find("_@_", System.Type.Missing,
                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                System.Type.Missing, System.Type.Missing);

            while (currentFind != null)
            {
                // Keep track of the first range you find.  
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done. 
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                currentFind.Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                // currentFind.Font.Bold = true;

                currentFind = range.FindNext(currentFind);
            }

            currentFind = null;
            firstFind = null;

            currentFind = range.Find("_@_", System.Type.Missing,
            Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
            System.Type.Missing, System.Type.Missing);

            while (currentFind != null)
            {
                // Keep track of the first range you find.  
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done. 
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                currentFind.Replace(What: "_@_", Replacement: "", LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: true, ReplaceFormat: false);

                currentFind = range.FindNext(currentFind);
            }





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

        public static string GenerateDocument(string filePath, List<QueryData> data, List<IndexData> index)
        {
            Excel.Application ExcelApp;
            Excel.Worksheet ExcelSheet;
            Excel.Workbook ExcelWorkbook;
            Excel.Workbooks ExcelWorkbooks;
            Excel.Range ExcelRange;
            int rowsCount;
            int colsCount;

            try
            {
                if (String.IsNullOrEmpty(filePath))
                {
                    filePath = Path.Combine(filePath + "\\" + data[0].ProjectName + "_Validation_" + DateTime.Now.ToShortDateString());
                    //if (Environment.OSVersion.Version.Major >= 6)
                    //{
                    //    filePath = Directory.GetParent(filePath).FullName;
                    //}
                }
                var filename = data[0].ProjectName + "_Validation_" + DateTime.Now.ToShortDateString() + ".xlsx";

                var indexDocument = index;

                ExcelApp = CreateExcelObj();
                ExcelWorkbooks = ExcelApp.Workbooks;
                ExcelApp.ScreenUpdating = false;
                ExcelApp.DisplayAlerts = false;
                ExcelWorkbook = ExcelWorkbooks.Add();

                data = SortData(data);

                foreach (var itemData in data)
                {
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
                    FormatHighlitedCells(rng);

                }

                //Adding Index
                object[,] dataIndex = new object[indexDocument.Count, 3];
                for (int j = 0; j < indexDocument.Count; j++)
                {
                    dataIndex[j, 0] = indexDocument[j].NameList;
                    dataIndex[j, 1] = indexDocument[j].ValidationRule;
                    dataIndex[j, 2] = indexDocument[j].Description;
                }
                ExcelSheet = ExcelWorkbook.Sheets.Add();
                ExcelSheet.Name = "Index";
                FormatSheet(ExcelSheet, data[0]);

                ExcelSheet.Cells[1, 1] = "Проект:";
                ExcelSheet.Cells[1, 2] = data[0].ProjectName;
                ExcelSheet.Cells[2, 1] = "Описание:";
                ExcelSheet.Range[ExcelSheet.Cells[2, 1], ExcelSheet.Cells[2, 3]].Merge();
                FormatDescription(ExcelSheet.Range[ExcelSheet.Cells[1, 1], ExcelSheet.Cells[2, 2]]);
                ExcelSheet.Cells[4, 1] = "Имя листа";
                ExcelSheet.Cells[4, 2] = "Правило валидации";
                ExcelSheet.Cells[4, 3] = "Описание правила для поиска ошибок";

                Excel.Range rngIndex = ExcelSheet.Range[ExcelSheet.Cells[5, 1], ExcelSheet.Cells[4 + indexDocument.Count, 3]];
                rngIndex.Value = dataIndex;
                FormatDataArea(ExcelSheet.Range[ExcelSheet.Cells[4, 1], ExcelSheet.Cells[4 + indexDocument.Count, 3]], "index");
                ExcelSheet.Range[ExcelSheet.Cells[5, 3], ExcelSheet.Cells[4 + indexDocument.Count, 3]].ColumnWidth = 90;
                ExcelSheet.Range[ExcelSheet.Cells[5, 3], ExcelSheet.Cells[4 + indexDocument.Count, 3]].WrapText = true;

                //End Creating Index

                ExcelWorkbook.SaveAs();
                ExcelWorkbook.SaveAs(filePath + "\\" + filename, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);


                while (Marshal.ReleaseComObject(ExcelWorkbook) > 0)
                { }
                while (Marshal.ReleaseComObject(ExcelWorkbooks) > 0)
                { }


                ExcelApp.Quit();

                while (Marshal.ReleaseComObject(ExcelApp) > 0)
                { }
                MessageBox.Show("Document created successfully !");
                return filePath + "\\" + filename;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Data + "\r\n" + ex.Message + "\r\n" + ex.Source + "\r\n" + ex.InnerException + "\r\n" +
                                ex.StackTrace);
                return null;
            }
            finally
            {

                GC.Collect();

            }
        }

        public static string GenerateDocument2(string filePath, List<QueryData> data, List<IndexData> index)
        {
            Excel.Application ExcelApp;
            Excel.Worksheet ExcelSheet;
            Excel.Workbook ExcelWorkbook;
            Excel.Workbooks ExcelWorkbooks;
            Excel.Range ExcelRange;
            int rowsCount;
            int colsCount;

            try
            {
                if (String.IsNullOrEmpty(filePath))
                {
                    filePath =
                        Path.Combine(filePath + "\\" + data[0].ProjectName + "_Export_" + DateTime.Now.ToShortDateString());
                    //if (Environment.OSVersion.Version.Major >= 6)
                    //{
                    //    filePath = Directory.GetParent(filePath).FullName;
                    //}
                }
                var filename = data[0].ProjectName + "_Export_" + DateTime.Now.ToShortDateString() + ".xlsx";

                var indexDocument = index;

                ExcelApp = CreateExcelObj();
                ExcelWorkbooks = ExcelApp.Workbooks;
                ExcelApp.ScreenUpdating = false;
                ExcelApp.DisplayAlerts = false;
                ExcelWorkbook = ExcelWorkbooks.Add();

                data = SortData(data);

                foreach (var itemData in data)
                {
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
                object[,] dataIndex = new object[indexDocument.Count, 3];
                for (int j = 0; j < indexDocument.Count; j++)
                {
                    dataIndex[j, 0] = indexDocument[j].NameList;
                    dataIndex[j, 1] = indexDocument[j].ValidationRule;
                    dataIndex[j, 2] = indexDocument[j].Description;
                }
                ExcelSheet = ExcelWorkbook.Sheets.Add();
                ExcelSheet.Name = "Index";
                FormatSheet(ExcelSheet, data[0]);

                ExcelSheet.Cells[1, 1] = "Проект:";
                ExcelSheet.Cells[1, 2] = data[0].ProjectName;
                ExcelSheet.Cells[2, 1] = "Описание:";
                ExcelSheet.Range[ExcelSheet.Cells[2, 1], ExcelSheet.Cells[2, 3]].Merge();
                FormatDescription(ExcelSheet.Range[ExcelSheet.Cells[1, 1], ExcelSheet.Cells[2, 2]]);
                ExcelSheet.Cells[4, 1] = "Имя листа";
                ExcelSheet.Cells[4, 2] = "Правило валидации";
                ExcelSheet.Cells[4, 3] = "Описание правила для поиска ошибок";

                Excel.Range rngIndex =
                    ExcelSheet.Range[ExcelSheet.Cells[5, 1], ExcelSheet.Cells[4 + indexDocument.Count, 3]];
                rngIndex.Value = dataIndex;
                FormatDataArea(ExcelSheet.Range[ExcelSheet.Cells[4, 1], ExcelSheet.Cells[4 + indexDocument.Count, 3]],
                    "index");
                ExcelSheet.Range[ExcelSheet.Cells[5, 3], ExcelSheet.Cells[4 + indexDocument.Count, 3]].ColumnWidth = 90;
                ExcelSheet.Range[ExcelSheet.Cells[5, 3], ExcelSheet.Cells[4 + indexDocument.Count, 3]].WrapText = true;

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
                MessageBox.Show("Document created successfully !");
                return filePath + "\\" + filename;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Data + "\r\n" + ex.Message + "\r\n" + ex.Source + "\r\n" + ex.InnerException + "\r\n" +
                                ex.StackTrace);
                return null;
            }
            finally
            {

                GC.Collect();

            }
        }
    }
}
