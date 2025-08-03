using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using System;
using System.IO;

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
                WordTable headerTable = document.Header.Default.AddTable(1, 1);
                headerTable.Rows[0].Cells[0].Paragraphs[0].Text = "H1";
                document.Footer.Default.AddParagraph("Example Footer");
                WordTable footerTable = document.Footer.Default.AddTable(1, 1);
                footerTable.Rows[0].Cells[0].Paragraphs[0].Text = "F1";

                WordParagraph heading = document.AddParagraph("Sample Heading");
                heading.Style = WordParagraphStyles.Heading1;

                WordParagraph formatted = document.AddParagraph("Bold Italic Underlined Centered");
                formatted.Bold = true;
                formatted.Italic = true;
                formatted.Underline = UnderlineValues.Single;
                formatted.ParagraphAlignment = JustificationValues.Center;

                WordList list = document.AddList(WordListStyle.ArticleSections);
                list.AddItem("First Item");
                list.AddItem("Second Item");

                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "B1";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "A2";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";
                WordTable nested = table.Rows[0].Cells[0].AddTable(1, 1);
                nested.Rows[0].Cells[0].Paragraphs[0].Text = "N1";

                document.AddParagraph().AddImage(imagePath, 50, 50);

                document.Save();
                PdfSaveOptions options = new PdfSaveOptions {
                    MarginLeft = 2,
                    MarginTop = 2,
                    MarginRight = 2,
                    MarginBottom = 2
                };
                document.SaveAsPdf(pdfPath, options);
            }
        }
    }
}