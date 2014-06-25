using Microsoft.Office.Interop.Excel;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes.Charts;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HeaderFooter = MigraDoc.DocumentObjectModel.HeaderFooter;
using Style = MigraDoc.DocumentObjectModel.Style;


namespace ExportValidation.Common
{
    public static class PDFGeneration
    {

        /// <summary>
        /// Defines the styles used in the document.
        /// </summary>
        private static void DefineStyles(Document document)
        {
            // Get the predefined style Normal.
            Style style = document.Styles["Normal"];
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Times New Roman";

            // Heading1 to Heading9 are predefined styles with an outline level. An outline level
            // other than OutlineLevel.BodyText automatically creates the outline (or bookmarks) 
            // in PDF.

            style = document.Styles["Heading1"];
            style.Font.Name = "Tahoma";
            style.Font.Size = 14;
            style.Font.Bold = true;
            style.Font.Color = Colors.DarkBlue;
            style.ParagraphFormat.PageBreakBefore = true;
            style.ParagraphFormat.SpaceAfter = 6;

            style = document.Styles["Heading2"];
            style.Font.Size = 12;
            style.Font.Bold = true;
            style.ParagraphFormat.PageBreakBefore = false;
            style.ParagraphFormat.SpaceBefore = 6;
            style.ParagraphFormat.SpaceAfter = 6;

            style = document.Styles["Heading3"];
            style.Font.Size = 10;
            style.Font.Bold = true;
            style.Font.Italic = true;
            style.ParagraphFormat.SpaceBefore = 6;
            style.ParagraphFormat.SpaceAfter = 3;

            style = document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);

            style = document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);

            // Create a new style called TextBox based on style Normal
            style = document.Styles.AddStyle("TextBox", "Normal");
            style.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
            style.ParagraphFormat.Borders.Width = 2.5;
            style.ParagraphFormat.Borders.Distance = "3pt";
            style.ParagraphFormat.Shading.Color = Colors.SkyBlue;

            // Create a new style called TOC based on style Normal
            style = document.Styles.AddStyle("TOC", "Normal");
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right, TabLeader.Dots);
            style.ParagraphFormat.Font.Color = Colors.Blue;
        }

        /// <summary>
        /// Defines the cover page.
        /// </summary>
        private static void DefineCover(Document document, string projectName)
        {
            Section section = document.AddSection();

            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.SpaceAfter = "3cm";

            //Image image = section.AddImage("../../images/Logo landscape.png");
            //image.Width = "10cm";

            paragraph = section.AddParagraph("Валидация данных по проекту " + projectName );
            paragraph.Format.Font.Size = 16;
            paragraph.Format.Font.Color = Colors.DarkRed;
            paragraph.Format.SpaceBefore = "8cm";
            paragraph.Format.SpaceAfter = "3cm";

            paragraph = section.AddParagraph("Дата создания: ");
            paragraph.AddDateField();
        }
        /// <summary>
        /// Defines the cover page.
        /// </summary>
        private static void DefineTableOfContents(Document document, List<QueryData> data )
        {
            Section section = document.LastSection;

            section.AddPageBreak();
            var paragraph = section.AddParagraph("Содержание:", "Heading1");
            
            paragraph.Format.Font.Size = 14;
            paragraph.Format.Font.Bold = true;
            paragraph.Format.SpaceAfter = 24;
            paragraph.Format.OutlineLevel = OutlineLevel.Level1;
        
            foreach (var item in data)
            {
                paragraph = section.AddParagraph();
                paragraph.Style = "TOC";
                var hyperlink = paragraph.AddHyperlink((item.NameList + " - " + item.ValidationRule).ToString());
                hyperlink.AddText(item.NameList + " - " + item.ValidationRule + "\t");
                hyperlink.AddPageRefField(item.NameList + " - " + item.ValidationRule);
            }
        }
        /// <summary>
        /// Defines page setup, headers, and footers.
        /// </summary>
        private static void DefineContentSection(Document document, List<QueryData> data )
        {
            Section section = document.AddSection();
            //section.PageSetup.OddAndEvenPagesHeaderFooter = true;
            section.PageSetup.StartingNumber = 1;

            HeaderFooter header = section.Headers.Primary;
            header.AddParagraph(data[0].ProjectName + " - Отчет проверки данных от " + DateTime.Now.ToShortDateString() );
            header.Format.Alignment = ParagraphAlignment.Center;
            // Create a paragraph with centered page number. See definition of style "Footer".
            Paragraph paragraph = new Paragraph();
            paragraph.AddTab();
            paragraph.AddPageField();
            paragraph.AddText(" из ");
            paragraph.AddNumPagesField();
            // Add paragraph to footer for odd pages.
            section.Footers.Primary.Add(paragraph);
            // Add clone of paragraph to footer for odd pages. Cloning is necessary because an object must
            // not belong to more than one other object. If you forget cloning an exception is thrown.
            section.Footers.EvenPage.Add(paragraph.Clone());
        }
        private static void DefineParagraphs(Document document)
        {
            Paragraph paragraph = document.LastSection.AddParagraph("Paragraph Layout Overview", "Heading1");
            paragraph.AddBookmark("Paragraphs");

            DemonstrateAlignment(document);
            DemonstrateIndent(document);
            DemonstrateFormattedText(document);
            DemonstrateBordersAndShading(document);
        }

        private static void DemonstrateAlignment(Document document)
        {
            document.LastSection.AddParagraph("Alignment", "Heading2");

            document.LastSection.AddParagraph("Left Aligned", "Heading3");

            Paragraph paragraph = document.LastSection.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Left;
            paragraph.AddText("лалоплодлога ыщгашг пыщ  щпшг ыщ ыщшывпшщыгпщыг ыпщ ыпшгыщпгыщпыщшпг ");

            document.LastSection.AddParagraph("Right Aligned", "Heading3");

            paragraph = document.LastSection.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Right;
            paragraph.AddText("лалоплодлога ыщгашг пыщ  щпшг ыщ ыщшывпшщыгпщыг ыпщ ыпшгыщпгыщпыщшпг ");

            document.LastSection.AddParagraph("Centered", "Heading3");

            paragraph = document.LastSection.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddText("лалоплодлога ыщгашг пыщ  щпшг ыщ ыщшывпшщыгпщыг ыпщ ыпшгыщпгыщпыщшпг ");

            document.LastSection.AddParagraph("Justified", "Heading3");

            paragraph = document.LastSection.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Justify;
            paragraph.AddText("лалоплодлога ыщгашг пыщ  щпшг ыщ ыщшывпшщыгпщыг ыпщ ыпшгыщпгыщпыщшпг ");
        }

        private static void DemonstrateIndent(Document document)
        {
            document.LastSection.AddParagraph("Indent", "Heading2");

            document.LastSection.AddParagraph("Left Indent", "Heading3");

            Paragraph paragraph = document.LastSection.AddParagraph();
            paragraph.Format.LeftIndent = "2cm";
            paragraph.AddText("лалоплодлога ыщгашг пыщ  щпшг ыщ ыщшывпшщыгпщыг ыпщ ыпшгыщпгыщпыщшпг ");

            document.LastSection.AddParagraph("Right Indent", "Heading3");

            paragraph = document.LastSection.AddParagraph();
            paragraph.Format.RightIndent = "1in";
            paragraph.AddText("лалоплодлога ыщгашг пыщ  щпшг ыщ ыщшывпшщыгпщыг ыпщ ыпшгыщпгыщпыщшпг ");

            document.LastSection.AddParagraph("First Line Indent", "Heading3");

            paragraph = document.LastSection.AddParagraph();
            paragraph.Format.FirstLineIndent = "12mm";
            paragraph.AddText("лалоплодлога ыщгашг пыщ  щпшг ыщ ыщшывпшщыгпщыг ыпщ ыпшгыщпгыщпыщшпг ");

            document.LastSection.AddParagraph("First Line Negative Indent", "Heading3");

            paragraph = document.LastSection.AddParagraph();
            paragraph.Format.LeftIndent = "1.5cm";
            paragraph.Format.FirstLineIndent = "-1.5cm";
            paragraph.AddText("лалоплодлога ыщгашг пыщ  щпшг ыщ ыщшывпшщыгпщыг ыпщ ыпшгыщпгыщпыщшпг ");
        }
        
        private static void DefineTables(Document document, List<QueryData> data )
        {
            foreach (var item in data)
            {
                Paragraph paragraph = document.LastSection.AddParagraph("Правило: "+item.NameList+" - " + item.ValidationRule, "Heading1");
                paragraph.AddBookmark(item.NameList + " - " + item.ValidationRule);

                SimpleTable(document, item);
               // DemonstrateAlignmentTable(document);
               // DemonstrateCellMerge(document);   
            }
        }

        private static void SimpleTable(Document document, QueryData data)
        {
            document.LastSection.AddParagraph("Описание: " + data.Description, "Heading2");

            Table table = new Table();
            table.Borders.Width = 0.75;

            Row row ;
            Cell cell;
            int i = 0;
            int cntr = 0;
            

            foreach (var strFieldName in data.FieldsName)
            {
            var column = table.AddColumn();
                column.Format.Alignment = ParagraphAlignment.Center;
                column.Width = (document.DefaultPageSetup.PageWidth - document.DefaultPageSetup.RightMargin - document.DefaultPageSetup.LeftMargin)/data.FieldsName.Count;
             

            }

            row = table.AddRow();
            row.Shading.Color = Colors.PaleGoldenrod;
            row.HeadingFormat = true;

            foreach (var strFieldName in data.FieldsName)
            {

                cell = row.Cells[i];
                cell.AddParagraph(strFieldName);
                i++;
            }
            for ( i = 0; i < data.Data.Rows.Count; i++)
            {
                cntr++;

                row = table.AddRow();
                if (cntr % 2 == 0)
                {
                    row.Shading.Color = Colors.LightGray;
                }
                for (int j = 0; j < data.FieldsName.Count; j++)
                {

                    cell = row.Cells[j];
                    cell.AddParagraph(data.Data.Rows[i][j].ToString());
                    cell.Format.Font.Size = 8;

                }
            }

    
       table.SetEdge(0, 0, data.FieldsName.Count, data.Data.Rows.Count, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);
   
            document.LastSection.Add(table);
        }

        private static void DemonstrateAlignmentTable(Document document)
        {
            document.LastSection.AddParagraph("Cell Alignment", "Heading2");

            Table table = document.LastSection.AddTable();
            table.Borders.Visible = true;
            table.Format.Shading.Color = Colors.LavenderBlush;
            table.Shading.Color = Colors.Salmon;
            table.TopPadding = 5;
            table.BottomPadding = 5;

            Column column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Left;

            column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Right;

            table.Rows.Height = 35;

            Row row = table.AddRow();
            row.VerticalAlignment = VerticalAlignment.Top;
            row.Cells[0].AddParagraph("Text");
            row.Cells[1].AddParagraph("Text");
            row.Cells[2].AddParagraph("Text");

            row = table.AddRow();
            row.VerticalAlignment = VerticalAlignment.Center;
            row.Cells[0].AddParagraph("Text");
            row.Cells[1].AddParagraph("Text");
            row.Cells[2].AddParagraph("Text");

            row = table.AddRow();
            row.VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("Text");
            row.Cells[1].AddParagraph("Text");
            row.Cells[2].AddParagraph("Text");
        }

        private static void DemonstrateCellMerge(Document document)
        {
            document.LastSection.AddParagraph("Cell Merge", "Heading2");

            Table table = document.LastSection.AddTable();
            table.Borders.Visible = true;
            table.TopPadding = 5;
            table.BottomPadding = 5;

            Column column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Left;

            column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Right;

            table.Rows.Height = 35;

            Row row = table.AddRow();
            row.Cells[0].AddParagraph("Merge Right");
            row.Cells[0].MergeRight = 1;

            row = table.AddRow();
            row.VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].MergeDown = 1;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("Merge Down");

            table.AddRow();
        }

        private static void DemonstrateFormattedText(Document document)
        {
            document.LastSection.AddParagraph("Formatted Text", "Heading2");

            //document.LastSection.AddParagraph("Left Aligned", "Heading3");

            Paragraph paragraph = document.LastSection.AddParagraph();
            paragraph.AddText("Text can be formatted ");
            paragraph.AddFormattedText("bold", TextFormat.Bold);
            paragraph.AddText(", ");
            paragraph.AddFormattedText("italic", TextFormat.Italic);
            paragraph.AddText(", or ");
            paragraph.AddFormattedText("bold & italic", TextFormat.Bold | TextFormat.Italic);
            paragraph.AddText(".");
            paragraph.AddLineBreak();
            paragraph.AddText("You can set the ");
            FormattedText formattedText = paragraph.AddFormattedText("size ");
            formattedText.Size = 15;
            paragraph.AddText("the ");
            formattedText = paragraph.AddFormattedText("color ");
            formattedText.Color = Colors.Firebrick;
            paragraph.AddText("the ");
            formattedText = paragraph.AddFormattedText("font", new MigraDoc.DocumentObjectModel.Font("Verdana"));
            paragraph.AddText(".");
            paragraph.AddLineBreak();
            paragraph.AddText("You can set the ");
            formattedText = paragraph.AddFormattedText("subscript");
            formattedText.Subscript = true;
            paragraph.AddText(" or ");
            formattedText = paragraph.AddFormattedText("superscript");
            formattedText.Superscript = true;
            paragraph.AddText(".");
        }

        static void DemonstrateBordersAndShading(Document document)
        {
            document.LastSection.AddPageBreak();
            document.LastSection.AddParagraph("Borders and Shading", "Heading2");

            document.LastSection.AddParagraph("Border around Paragraph", "Heading3");

            Paragraph paragraph = document.LastSection.AddParagraph();
            paragraph.Format.Borders.Width = 2.5;
            paragraph.Format.Borders.Color = Colors.Navy;
            paragraph.Format.Borders.Distance = 3;
            paragraph.AddText("лалоплодлога ыщгашг пыщ  щпшг ыщ ыщшывпшщыгпщыг ыпщ ыпшгыщпгыщпыщшпг ");

            document.LastSection.AddParagraph("Shading", "Heading3");

            paragraph = document.LastSection.AddParagraph();
            paragraph.Format.Shading.Color = Colors.LightCoral;
            paragraph.AddText("лалоплодлога ыщгашг пыщ  щпшг ыщ ыщшывпшщыгпщыг ыпщ ыпшгыщпгыщпыщшпг ");

            document.LastSection.AddParagraph("Borders & Shading", "Heading3");

            paragraph = document.LastSection.AddParagraph();
            paragraph.Style = "TextBox";
            paragraph.AddText("лалоплодлога ыщгашг пыщ  щпшг ыщ ыщшывпшщыгпщыг ыпщ ыпшгыщпгыщпыщшпг ");
        }

        public static Document CreateDocument(List<QueryData> data )
        {
            // Create a new MigraDoc document
            Document document = new Document();
            DefineStyles(document);
            DefineCover(document, data[0].ProjectName);
            DefineTableOfContents(document, data);
            DefineContentSection(document, data);
        //    DefineParagraphs(document);
            DefineTables(document, data);
            //DefineCharts(document);
            return document;
        }


        public static void GenerateDocument(string filePath, List<QueryData> data)
        {
            Document document = CreateDocument(data);
            MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToFile(document, "MigraDoc.mdddl");

            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true, PdfSharp.Pdf.PdfFontEmbedding.Always);
            renderer.Document = document;
            renderer.RenderDocument();

            // Save the document...
            string filename = data[0].ProjectName + "_Validation_" + DateTime.Now.ToShortDateString() + ".pdf";
            renderer.PdfDocument.Save(filename);
            // ...and start a viewer.
            Process.Start(filename);


        }






    }
}
