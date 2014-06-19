using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
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

        private static void FormatDataArea(Excel.Range range, QueryData data)
        {
            range.Worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, range, System.Type.Missing,
                Excel.XlYesNoGuess.xlYes, System.Type.Missing).Name = data.ValidationRule;
            range.Select();
            range.Worksheet.ListObjects[data.ValidationRule].TableStyle = "TableStyleMedium2";
            range.EntireColumn.AutoFit();
        }

        private static void FormatSheet(Excel.Worksheet sheet, QueryData obj )
        {
            sheet.PageSetup.PrintGridlines = false;
            sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
            sheet.PageSetup.RightFooter = "Дата: &DD Стр &PP из &NN";
            sheet.PageSetup.RightHeader = obj.ProjectName + " - " + obj.ValidationRule;
            sheet.PageSetup.Zoom = false;
            sheet.PageSetup.LeftHeader = "e-CRF";
            sheet.PageSetup.TopMargin = 50;
            sheet.PageSetup.BottomMargin = 50;
            sheet.PageSetup.HeaderMargin = 20;
            sheet.PageSetup.FooterMargin = 20;
            sheet.PageSetup.RightMargin = 10;
            sheet.PageSetup.LeftMargin = 50;
            sheet.PageSetup.Order = Excel.XlOrder.xlOverThenDown;
            sheet.Columns.EntireColumn.AutoFit();
        }

        public static string  GenerateDocument(string filePath, List<QueryData> data)
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
                var filename = data[0].ProjectName + "_Validation_"+DateTime.Now.ToShortDateString() + ".xlsx";

                ExcelApp = CreateExcelObj();
                ExcelWorkbooks = ExcelApp.Workbooks;
                ExcelApp.ScreenUpdating = false;
                ExcelApp.DisplayAlerts = false;
                ExcelWorkbook = ExcelWorkbooks.Add();

                foreach (var itemData in data)
                {
                    ExcelSheet = ExcelWorkbook.Sheets.Add();
                    ExcelSheet.Name = itemData.NameList;
                    FormatSheet(ExcelSheet, itemData);
                    rowsCount = itemData.Data.Rows.Count;
                    colsCount = itemData.Data.Columns.Count;
                    
                    ExcelSheet.Cells[1, 1] = "Проект:";
                    ExcelSheet.Cells[1, 2] = itemData.ProjectName;
                    ExcelSheet.Range[ExcelSheet.Cells[1,2],ExcelSheet.Cells[1,colsCount]].Merge();
                    ExcelSheet.Cells[2, 1] = "Правило:";
                    ExcelSheet.Cells[2, 2] = itemData.ValidationRule;
                    ExcelSheet.Range[ExcelSheet.Cells[2, 2], ExcelSheet.Cells[2, colsCount]].Merge();
                    ExcelSheet.Cells[3, 1] = "Описание:";
                    ExcelSheet.Cells[3, 2] = itemData.Description;
                    ExcelSheet.Range[ExcelSheet.Cells[3, 2], ExcelSheet.Cells[3, colsCount]].Merge();
                      

                    for (int i = 1; i <= itemData.FieldsName.Count; i++)
                    {
                        ExcelSheet.Cells[5, i] = itemData.FieldsName[i - 1];
                    }

                    for (int i = 0; i < itemData.Data.Rows.Count; i++)
                    {
                        for (int j = 0; j < itemData.FieldsName.Count; j++)
                        {
                            ExcelSheet.Cells[6 + i, 1 + j] = itemData.Data.Rows[i][j].ToString();
                        }
                    }
                    FormatDataArea(ExcelSheet.Range[ExcelSheet.Cells[5,1],ExcelSheet.Cells[5+rowsCount,colsCount]],itemData);
                    FormatDescription(ExcelSheet.Range[ExcelSheet.Cells[1, 1], ExcelSheet.Cells[3, 2]]);
                    ExcelSheet.Range[ExcelSheet.Cells[3, 2], ExcelSheet.Cells[3, colsCount]].WrapText = true;
                    ExcelSheet.Range[ExcelSheet.Cells[3, 2], ExcelSheet.Cells[3, colsCount]].RowHeight = 60;

                }
                   
                    
                ExcelWorkbook.SaveAs();
                ExcelWorkbook.SaveAs(filePath + "\\" + filename, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);


                while (Marshal.ReleaseComObject(ExcelWorkbook) > 0)
                { }
                while (Marshal.ReleaseComObject(ExcelWorkbooks) > 0)
                { }


                ExcelApp.Quit();

                while (Marshal.ReleaseComObject(ExcelApp) > 0)
                { }

                return filePath + "\\" + filename;

            }
            catch (Exception ex)
            {
                return ex.Data + "\r\n" + ex.Message + "\r\n" + ex.Source + "\r\n" + ex.InnerException + "\r\n" +
                       ex.StackTrace;
            }
            finally
            {

                GC.Collect();

            }
        }
    }
}
