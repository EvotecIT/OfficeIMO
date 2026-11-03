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
                    page.Margin(OfficeLength.Points(36), OfficeLength.Points(42), OfficeLength.Points(36), OfficeLength.Points(42));
                    page.Header(h => h.Alignment(OfficeAlignment.Left).Text("OfficeIMO.Pdf composition"));

                    page.Content(c => c
                        .Column(column => {
                            column.Item().H1("Compose DSL");
                            column.Item().AddParagraph(p => p.AddText("A compact report composed from pages, columns and styled text."));
                            column.Item().AddPanelParagraph(
                                p => {
                                    p.AddText("Report structure").SetBold();
                                    p.AddLineBreak();
                                    p.AddText("Each page combines text, panels and lists with a shared header and page count.");
                                },
                                new PdfPanelStyle {
                                    Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                                    PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                                    PaddingY = global::OfficeIMO.OfficeLength.Points(7)
                                });
                            column.Item().PageBreak();

                            column.Item().H2("Operational Notes");
                            column.Item().AddList(new[] {
                                "Each section starts on its own page.",
                                "The same theme carries through all three pages.",
                                "Page numbers include the complete report."
                            }, color: OfficeColor.FromRgb(55, 65, 81));
                            column.Item().HR(OfficeLength.Points(0.8), OfficeColor.FromRgb(183, 194, 207), OfficeLength.Points(8), OfficeLength.Points(8));
                            column.Item().AddParagraph(p => {
                                p.AddText("Status: ");
                                p.AddText("ready for review. ").SetBold();
                                p.AddText("All sections are complete.").SetColor(OfficeColor.FromRgb(20, 90, 180));
                            });
                            column.Item().PageBreak();

                            column.Item().H2("Color Sections");
                            column.Item().AddParagraph(p => p.AddText("Critical items should remain readable when colored inline.").SetColor(global::OfficeIMO.OfficeColor.FromRgb(185, 28, 28)));
                            column.Item().AddParagraph(p => p.AddText("Healthy items should use calm green without overpowering the page.").SetColor(global::OfficeIMO.OfficeColor.FromRgb(22, 101, 52)));
                            column.Item().AddParagraph(p => p.AddText("Informational items should keep good contrast on white backgrounds.").SetColor(global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180)));
                            column.Item().AddParagraph(p => p.AddText("End of compose DSL sample."), OfficeAlignment.Right, global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80));
                        }));

                    page.Footer(f => f.Alignment(OfficeAlignment.Center).Text(t => t
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
