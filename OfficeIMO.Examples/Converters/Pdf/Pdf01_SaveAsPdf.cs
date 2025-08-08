using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Pdf01_SaveAsPdf {
        public static void Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating Word document and exporting to PDF");
            
            using var doc = WordDocument.Create();
            
            // Add content
            doc.AddParagraph("PDF Export Example").Style = WordParagraphStyles.Heading1;
            doc.AddParagraph("This document demonstrates PDF export functionality.");
            
            doc.AddParagraph("Features").Style = WordParagraphStyles.Heading2;
            
            var paragraph = doc.AddParagraph("Text can be ");
            paragraph.AddText("bold").Bold = true;
            paragraph.AddText(", ");
            paragraph.AddText("italic").Italic = true;
            paragraph.AddText(", or ");
            var bothText = paragraph.AddText("both");
            bothText.Bold = true;
            bothText.Italic = true;
            
            var list = doc.AddList(WordListStyle.Heading1ai);
            list.AddItem("First numbered item");
            list.AddItem("Second numbered item");
            list.AddItem("Third numbered item");
            
            doc.AddParagraph("Table Example").Style = WordParagraphStyles.Heading3;
            
            var table = doc.AddTable(3, 3, WordTableStyle.GridTable4Accent1);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Column A";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Column B";
            table.Rows[0].Cells[2].Paragraphs[0].Text = "Column C";
            
            for (int i = 1; i < 3; i++) {
                for (int j = 0; j < 3; j++) {
                    table.Rows[i].Cells[j].Paragraphs[0].Text = $"Data {i},{j}";
                }
            }
            
            // Save as PDF
            string outputPath = Path.Combine(folderPath, "SaveAsPdf.pdf");
            doc.SaveAsPdf(outputPath, new PdfSaveOptions {
                Orientation = PdfPageOrientation.Portrait,
                Margin = 2,
                MarginUnit = QuestPDF.Infrastructure.Unit.Centimetre,
                MarginTop = 3,
                MarginTopUnit = QuestPDF.Infrastructure.Unit.Centimetre,
                MarginLeft = 1,
                MarginLeftUnit = QuestPDF.Infrastructure.Unit.Centimetre
            });
            
            Console.WriteLine($"✓ Created: {outputPath}");
            
            // Also save as DOCX for comparison
            string docxPath = Path.Combine(folderPath, "SaveAsPdf_Source.docx");
            doc.Save(docxPath);
            
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
            }
        }
    }
}