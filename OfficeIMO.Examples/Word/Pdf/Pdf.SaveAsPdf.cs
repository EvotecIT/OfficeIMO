using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveAsPdf(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document and exporting to PDF");
            string docPath = Path.Combine(folderPath, "ExportToPdf.docx");
            string pdfPath = Path.Combine(folderPath, "ExportToPdf.pdf");
            string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "Images", "EvotecLogo.png");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddHeadersAndFooters();
                document.Header.Default.AddParagraph("Example Header");
                document.Footer.Default.AddParagraph("Example Footer");

                var heading = document.AddParagraph("Sample Heading");
                heading.Style = WordParagraphStyles.Heading1;

                var formatted = document.AddParagraph("Bold Italic Underlined Centered");
                formatted.Bold = true;
                formatted.Italic = true;
                formatted.Underline = UnderlineValues.Single;
                formatted.ParagraphAlignment = JustificationValues.Center;

                var list = document.AddList(WordListStyle.ArticleSections);
                list.AddItem("First Item");
                list.AddItem("Second Item");

                var table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "B1";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "A2";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";

                document.AddParagraph().AddImage(imagePath, 50, 50);

                document.Save();
                document.SaveAsPdf(pdfPath);
            }
        }
    }
}
