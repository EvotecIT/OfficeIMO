using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class FlowDslDemo {
        public static void Example_Pdf_FlowDslDemo(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.FlowDslDemo.pdf");

            var options = new PdfCreateOptions();

            var doc = PdfDocument.Create(document => {
                document.Page(page => {
                    page.Size(PdfPageSizes.A5);
                    page.Margin(36, 42, 36, 42);
                    page.Header(h => h.AlignLeft().Text("OfficeIMO.Pdf composition"));

                    page.Content(c => c
                        .Column(column => {
                            column.Item().H1("Compose DSL");
                            column.Item().Paragraph(p => p.Text("A compact report composed from pages, columns and styled text."));
                            column.Item().PanelParagraph(
                                p => p
                                    .Bold("Report structure")
                                    .LineBreak()
                                    .Text("Each page combines text, panels and lists with a shared header and page count."),
                                new PdfPanelStyle {
                                    Background = PdfColor.FromRgb(248, 250, 252),
                                    BorderColor = PdfColor.FromRgb(183, 194, 207),
                                    PaddingX = 9,
                                    PaddingY = 7
                                });
                            column.Item().PageBreak();

                            column.Item().H2("Operational Notes");
                            column.Item().Bullets(new[] {
                                "Each section starts on its own page.",
                                "The same theme carries through all three pages.",
                                "Page numbers include the complete report."
                            }, color: PdfColor.FromRgb(55, 65, 81));
                            column.Item().HR(0.8, PdfColor.FromRgb(183, 194, 207), 8, 8);
                            column.Item().Paragraph(p => p
                                .Text("Status: ")
                                .Bold("ready for review. ")
                                .Color(PdfColor.FromRgb(20, 90, 180)).Text("All sections are complete."));
                            column.Item().PageBreak();

                            column.Item().H2("Color Sections");
                            column.Item().Paragraph(p => p.Color(PdfColor.FromRgb(185, 28, 28)).Text("Critical items should remain readable when colored inline."));
                            column.Item().Paragraph(p => p.Color(PdfColor.FromRgb(22, 101, 52)).Text("Healthy items should use calm green without overpowering the page."));
                            column.Item().Paragraph(p => p.Color(PdfColor.FromRgb(20, 90, 180)).Text("Informational items should keep good contrast on white backgrounds."));
                            column.Item().Paragraph(p => p.Text("End of compose DSL sample."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80));
                        }));

                    page.Footer(f => f.AlignCenter().Text(t => t
                        .Text("Page ")
                        .CurrentPage()
                        .Text(" of ")
                        .TotalPages()));
                });
            }, options);

            doc.Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
