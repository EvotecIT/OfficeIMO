using OfficeIMO.Word;

namespace OfficeIMO.Examples.Showcase;

/// <summary>Creates the browser conversion sample using its bundled Latin font family.</summary>
internal static class BrowserBusinessSample {
    internal static void Create(string outputDirectory) {
        Directory.CreateDirectory(outputDirectory);
        using WordDocument document = WordDocument.Create(Path.Combine(outputDirectory, "business-summary.docx"));
        document.BuiltinDocumentProperties.Title = "Monthly operations brief";
        document.BuiltinDocumentProperties.Creator = "OfficeIMO";
        document.Settings.FontFamily = "Carlito";
        document.Settings.FontFamilyHighAnsi = "Carlito";
        document.Settings.FontSize = 11;

        Paragraph("NORTHWIND OPERATIONS", 10, true);
        Paragraph("Monthly operations brief", 26, true);
        Paragraph("August 2026 | Prepared for the service review", 11);
        Paragraph("A clear view of delivery, service health, and the decisions for next month.", 12);

        Paragraph("At a glance", 16, true);
        Table(new[,] {
            { "Measure", "This month", "Status" },
            { "Requests completed", "248", "On plan" },
            { "Service availability", "99.95%", "Target met" },
            { "Customer satisfaction", "4.8 / 5", "Improving" }
        });

        Paragraph("What changed", 16, true);
        Paragraph("The team completed the reporting rollout and reduced the oldest support queue. " +
            "The new handover checklist gives each request a named owner and a clear next action.", 11);

        Paragraph("Next actions", 16, true);
        Table(new[,] {
            { "Action", "Owner", "Due" },
            { "Review the September forecast", "Finance", "10 September" },
            { "Confirm the recovery exercise", "Operations", "18 September" },
            { "Publish the service handbook", "Delivery", "25 September" }
        });
        Paragraph("Prepared from sample data. Generated with OfficeIMO.Word using the bundled Carlito font.", 9);
        document.Save();

        void Paragraph(string text, int size, bool bold = false) {
            WordParagraph paragraph = document.AddParagraph(text);
            paragraph.FontFamily = "Carlito";
            paragraph.FontSize = size;
            paragraph.Bold = bold;
            paragraph.Color = OfficeIMO.Drawing.OfficeColor.FromRgb(15, 23, 42);
            paragraph.LineSpacingBeforePoints = size == 16 ? 14 : 3;
            paragraph.LineSpacingAfterPoints = size == 26 ? 10 : 6;
        }

        void Table(string[,] values) {
            WordTable table = document.AddTable(values.GetLength(0), values.GetLength(1), WordTableStyle.TableGrid);
            for (int row = 0; row < values.GetLength(0); row++) {
                for (int column = 0; column < values.GetLength(1); column++) {
                    WordParagraph paragraph = table.Rows[row].Cells[column].Paragraphs[0];
                    paragraph.Text = values[row, column];
                    paragraph.FontFamily = "Carlito";
                    paragraph.FontSize = 11;
                    paragraph.Bold = row == 0;
                    paragraph.LineSpacingBeforePoints = 3;
                    paragraph.LineSpacingAfterPoints = 3;
                    table.Rows[row].Cells[column].MarginTopCentimeters = 0.08;
                    table.Rows[row].Cells[column].MarginBottomCentimeters = 0.08;
                    if (row == 0) {
                        table.Rows[row].Cells[column].ShadingFillColorHex = "E8EEF8";
                    }
                }
            }
        }
    }
}
