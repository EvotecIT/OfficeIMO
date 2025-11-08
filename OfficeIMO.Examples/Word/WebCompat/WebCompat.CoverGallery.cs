using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class WebCompat {
        public static void Example_CoverTemplates_Basic(string folderPath, bool openWord) {
            // A small set of templates to evaluate in Word Online / Google Docs
            var templates = new[] {
                CoverPageTemplate.IonLight,
                CoverPageTemplate.SideLine,
                CoverPageTemplate.Retrospect
            };

            foreach (var tpl in templates) {
                string filePath = System.IO.Path.Combine(folderPath, $"WebCompat-Cover-{tpl}.docx");
                Console.WriteLine("[*] Generating: " + filePath);
                using var doc = WordDocument.Create(filePath);
                doc.BuiltinDocumentProperties.Title = $"Cover: {tpl}";
                doc.ApplicationProperties.Company = "OfficeIMO";

                doc.AddCoverPage(tpl);
                doc.AddPageBreak();
                doc.AddParagraph("1. Executive Summary");
                var t = doc.AddTable(2, 2, WordTableStyle.TableGrid);
                t.WidthType = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Pct; t.Width = 5000;
                t.ColumnWidthType = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Pct; t.ColumnWidth = new() { 1500, 3500 };
                t.Rows[0].Cells[0].AddParagraph("Name", true);
                t.Rows[0].Cells[1].AddParagraph("Value", true);
                t.Rows[1].Cells[0].AddParagraph("Example", true);
                t.Rows[1].Cells[1].AddParagraph("Data", true);

                doc.Save(openWord);
            }
        }

        public static void Example_CoverWithConfidentialWatermark(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "WebCompat-Cover-Confidential.docx");
            Console.WriteLine("[*] Generating: " + filePath);
            using var doc = WordDocument.Create(filePath);
            doc.BuiltinDocumentProperties.Title = "Confidential Report";
            doc.AddCoverPage(CoverPageTemplate.IonDark);

            // Optional: add text watermark so we can observe how Online renders it
            try { doc.Sections[0].AddWatermark(WordWatermarkStyle.Text, "CONFIDENTIAL", scale: 1.2); } catch { }

            doc.AddPageBreak();
            doc.AddParagraph("1. Overview");
            var t = doc.AddTable(2, 2, WordTableStyle.TableGrid);
            t.ColumnWidthType = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Pct; t.ColumnWidth = new() { 500, 4500 };
            t.WidthType = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Pct; t.Width = 5000;
            t.Rows[0].Cells[0].AddParagraph("10%", true);
            t.Rows[0].Cells[1].AddParagraph("90%", true);
            t.Rows[1].Cells[0].AddParagraph("10%", true);
            t.Rows[1].Cells[1].AddParagraph("90%", true);
            doc.Save(openWord);
        }
    }
}

