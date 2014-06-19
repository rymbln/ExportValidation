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



namespace ExportValidation.Common
{
    public static class PDFGeneration
    {

        public static Document CreateDocument()
        {
            // Create a new MigraDoc document
            Document document = new Document();
            document.Info.Title = "Hello, MigraDoc";
            document.Info.Subject = "Demonstrates an excerpt of the capabilities of MigraDoc.";
            document.Info.Author = "Stefan Lange";
            document.AddSection();
            document.LastSection.AddParagraph("вылоплдыопловылдповылдоп", "Heading2");

            Table table = new Table();
            table.Borders.Width = 0.75;

            Column column = table.AddColumn(Unit.FromCentimeter(2));
            column.Format.Alignment = ParagraphAlignment.Center;

            table.AddColumn(Unit.FromCentimeter(5));

            Row row = table.AddRow();
            row.Shading.Color = Colors.PaleGoldenrod;
            Cell cell = row.Cells[0];
            cell.AddParagraph("Itemus");
            cell = row.Cells[1];
            cell.AddParagraph("Descriptum");

            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("1");
            cell = row.Cells[1];

            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("2");
            cell = row.Cells[1];


            table.SetEdge(0, 0, 2, 3, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);
    
     


            return document;
        }
      

        public static void GenerateDocument(string filePath, List<QueryData> data)
        {
            Document document = CreateDocument();
            MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToFile(document, "MigraDoc.mdddl");
            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true, PdfSharp.Pdf.PdfFontEmbedding.Always);
            renderer.Document = document;
            renderer.RenderDocument();
            string filename = "HelloMigraDoc.pdf";
            renderer.PdfDocument.Save(filename);
            Process.Start(filename);
        }
        //var doc = new Document();

        //PdfWriter.GetInstance(doc, new FileStream(filePath + @"\Document.pdf", FileMode.Create));
        //doc.Open();
        //QueryData dataItem;

        //string fontPath = Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\arial.ttf";
        //iTextSharp.text.FontFactory.Register(fontPath);
        //BaseFont baseFont = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        //iTextSharp.text.Font f = new iTextSharp.text.Font(baseFont, 12);
        //foreach (var d in data)
        //{



        //    dataItem = d;
        //    var intColumns = dataItem.FieldsName.Count;
        //    PdfPTable table = new PdfPTable(intColumns);

        //    PdfPCell cell = new PdfPCell(new Phrase(dataItem.ValidationRule, f));
        //    cell.BackgroundColor = new BaseColor(Color.Wheat);
        //    cell.Padding = 5;
        //    cell.Colspan = intColumns;
        //    cell.HorizontalAlignment = Element.ALIGN_CENTER;

        //    table.AddCell(cell);

        //    foreach (var field in dataItem.FieldsName)
        //    {
        //        table.AddCell(field);
        //    }

        //    for (int i = 0; i < dataItem.Data.Rows.Count; i++)
        //    {
        //        for (int j = 0; j < dataItem.FieldsName.Count; j++)
        //        {
        //            table.AddCell(dataItem.Data.Rows[i][j].ToString());
        //        }
        //    }
        //    doc.Add(table);
        //}

        //doc.Close();





    }
}
