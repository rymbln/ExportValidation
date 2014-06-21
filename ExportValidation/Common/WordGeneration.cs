using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace ExportValidation.Common
{
    public static class WordGeneration
    {
        private static Word.Application CreateWordObj()
        {
            object obj;
            obj = null;
            try
            {
                Word.Application objWord = new Word.Application();
                objWord.Visible = true;
                obj = objWord;
            }
            catch (Exception)
            {
                throw new Exception("Ошибка создания экземпляра MS Word");
            }
            return (obj as Word.Application);
        }


        public static string GenerateDocument(string filePath, List<QueryData> data)
        {
            object missing = System.Reflection.Missing.Value;
            int rowsCount;
            int colsCount;
            try
            {
                if (String.IsNullOrEmpty(filePath))
                {
                    filePath =
                        Path.Combine(filePath + "\\" + data[0].ProjectName + "_Validation_" +
                                     DateTime.Now.ToShortDateString());
                    //if (Environment.OSVersion.Version.Major >= 6)
                    //{
                    //    filePath = Directory.GetParent(filePath).FullName;
                    //}
                }
                var filename = data[0].ProjectName + "_Validation_" + DateTime.Now.ToShortDateString() + ".docx";

                var objWord = CreateWordObj();

                var doc = objWord.Documents.Add();

                //retrieve the first paragragh of the document
                Word.Paragraph paragraphBefore = doc.Paragraphs[1];
                Word.Range rngBefore1 = paragraphBefore.Range;

                //insert two paragraphs for table of content and title
                rngBefore1.InsertParagraphBefore();
                rngBefore1.InsertParagraphBefore();

                //retrieve the third paragragh for inserting page break later
                rngBefore1 = doc.Paragraphs[3].Range;
                // rngBefore1.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                //insert a page break
                rngBefore1.InsertBreak(Word.WdBreakType.wdPageBreak);

                Word.Paragraph paragraphAfter1 = doc.Paragraphs[1];
                Word.Paragraph paragraphAfter2 = doc.Paragraphs[2];


                /*======================================================================*/
                /* How to insert <Current Page> of <Total Pages> into the footer for a MS Word Document in VSTO*/

                // Open up the footer in the word document

                objWord.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;


                // Set current Paragraph Alignment to Center
                objWord.ActiveWindow.ActivePane.Selection.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                // Type in 'Page '
                objWord.ActiveWindow.Selection.TypeText("Страница ");

                // Add in current page field
                Object CurrentPage = Word.WdFieldType.wdFieldPage;
                objWord.ActiveWindow.Selection.Fields.Add(objWord.ActiveWindow.Selection.Range, ref CurrentPage, ref missing, ref missing);

                // Type in ' of '
                objWord.ActiveWindow.Selection.TypeText(" из ");

                // Add in total page field
                Object TotalPages = Word.WdFieldType.wdFieldNumPages;
                objWord.ActiveWindow.Selection.Fields.Add(objWord.ActiveWindow.Selection.Range, ref TotalPages, ref missing, ref missing);
                /*======================================================================*/

                foreach (var itemData in data)
                {
                   
                    //Add header
                    foreach (Word.Section section in doc.Sections)
                    {
                        //Get the header range and add the header details.
                        Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                        headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        headerRange.Font.ColorIndex = Word.WdColorIndex.wdGray50;
                        headerRange.Font.Size = 10;
                        headerRange.Text = itemData.ProjectName + " - Отчет проверки данных от " + DateTime.Now.ToShortDateString();
                    }
                    // Add page numbers in footer

                    //Add the footers into the document
                    foreach (Word.Section wordSection in doc.Sections)
                    {
                        //Get the footer range and add the footer details.
                        Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        footerRange.Font.ColorIndex = Word.WdColorIndex.wdGray50;
                        footerRange.Font.Size = 10;
                        footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        footerRange.Text = DateTime.Now.ToShortDateString();
                    }

                    //Add paragraph with Heading 1 style
                    Word.Paragraph para1 = doc.Content.Paragraphs.Add();
                    object styleHeading1 = "Заголовок 1";
                    para1.Range.set_Style(ref styleHeading1);
                    para1.Range.Text = "Правило: " + itemData.NameList + " - " + itemData.ValidationRule;
                    para1.Range.InsertParagraphAfter();

                    //Add paragraph with Heading 2 style
                    Word.Paragraph para2 = doc.Content.Paragraphs.Add();
                    object styleHeading2 = Word.WdBuiltinStyle.wdStyleHeading2;
                    para2.Range.set_Style(ref styleHeading2);
                    para2.set_Style(ref styleHeading2);
                    para2.Range.Text = "Описание: " + itemData.Description;
                    para2.Range.InsertParagraphAfter();

                    //Create a 5X5 table and insert some dummy record
                    Word.Paragraph para3 = doc.Content.Paragraphs.Add();
                    rowsCount = itemData.Data.Rows.Count;
                    colsCount = itemData.Data.Columns.Count;
                    Word.Table firstTable = doc.Tables.Add(para3.Range,rowsCount, colsCount);

                    firstTable.Borders.Enable = 1;


                    foreach (Word.Row row in firstTable.Rows)
                    {
                        foreach (Word.Cell cell in row.Cells)
                        {
                            //Header row
                            if (cell.RowIndex == 1)
                            {
                                cell.Range.Text = itemData.FieldsName[ cell.ColumnIndex-1];
                                cell.Range.Font.Bold = 1;
                                //other format properties goes here
                                cell.Range.Font.Name = "verdana";
                                cell.Range.Font.Size = 10;
                                //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                            
                                cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
                                //Center alignment for the Header cells
                                cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            }
                            //Data row
                            else
                            {
                                cell.Range.Text = itemData.Data.Rows[cell.RowIndex - 2][cell.ColumnIndex - 1].ToString();

                            }
                        }
                    }
                   
                    doc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }


                // Add Table Of Contents


                Word.Range rangeTOC = paragraphAfter1.Range;
                rangeTOC.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            
                object oTrueValue = true;

                Word.TableOfContents toc = doc.TablesOfContents.Add(rangeTOC);
                toc.Update();


                Word.Range rngTOC = toc.Range;
                rngTOC.Font.Size = 10;
                rngTOC.Font.Name = "Georgia";


                Word.Range rangeTOC2 = paragraphAfter2.Range;
                rangeTOC2.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                Word.TableOfContents toc2=doc.TablesOfContents.Add(rangeTOC2, ref missing, ref missing, ref missing, ref missing, ref missing,ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
toc2.Update();







                //Save the document
                doc.SaveAs(filePath + "\\" + filename);
                doc.Close();
                doc = null;
                objWord.Quit();
                objWord = null;
                MessageBox.Show("Document created successfully !");


                return filePath + "\\" + filename;
            }
            catch (Exception ex)
            {
                MessageBox.Show( ex.Data + "\r\n" + ex.Message + "\r\n" + ex.Source + "\r\n" + ex.InnerException + "\r\n" +ex.StackTrace);
                return null;
            }
            finally
            {

                GC.Collect();

            }
        }
    }
}
