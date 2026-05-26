using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class FlowDslDemo {
        public static void Example_Pdf_FlowDslDemo(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.FlowDslDemo.pdf");

            var options = new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8
            };

            var doc = PdfDoc.Create(options).Compose(document => {
                document.Page(page => {
                    page.Size(PageSizes.A5);
                    page.Margin(36, 42, 36, 42);
                    page.DefaultTextStyle(x => x
                        .Font(PdfStandardFont.Helvetica)
                        .FontSize(10)
                        .Color(PdfColor.FromRgb(31, 41, 55)));
                    page.Header(h => h.AlignLeft().Text("OfficeIMO.Pdf compose DSL gate"));

                    page.Content(c => c
                        .Column(column => {
                            column.Item().H1("Compose DSL");
                            column.Item().Paragraph(p => p.Text("A compact multi-page visual baseline for the OfficeIMO.Pdf composition API."));
                            column.Item().PanelParagraph(
                                p => p
                                    .Bold("What this protects")
                                    .LineBreak()
                                    .Text("Page settings, composed content, explicit page breaks, header/footer tokens, and rich text inside composed items."),
                                new PanelStyle {
                                    Background = PdfColor.FromRgb(248, 250, 252),
                                    BorderColor = PdfColor.FromRgb(183, 194, 207),
                                    PaddingX = 9,
                                    PaddingY = 7
                                });
                            column.Item().PageBreak();

                            column.Item().H2("Operational Notes");
                            column.Item().Bullets(new[] {
                                "Compose uses the same document engine as fluent blocks.",
                                "Visual gates should cover both public authoring styles.",
                                "Footer page totals must remain correct after explicit breaks."
                            }, color: PdfColor.FromRgb(55, 65, 81));
                            column.Item().HR(0.8, PdfColor.FromRgb(183, 194, 207), 8, 8);
                            column.Item().Paragraph(p => p
                                .Text("Status: ")
                                .Bold("ready for wrapper experiments ")
                                .Color(PdfColor.FromRgb(20, 90, 180)).Text("once visual quality keeps improving."));
                            column.Item().PageBreak();

                            column.Item().H2("Color Sections");
                            column.Item().Paragraph(p => p.Color(PdfColor.FromRgb(185, 28, 28)).Text("Critical items should remain readable when colored inline."));
                            column.Item().Paragraph(p => p.Color(PdfColor.FromRgb(22, 101, 52)).Text("Healthy items should use calm green without overpowering the page."));
                            column.Item().Paragraph(p => p.Color(PdfColor.FromRgb(20, 90, 180)).Text("Informational items should keep good contrast on white backgrounds."));
                            column.Item().Paragraph(p => p.Text("End of compose DSL sample."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80));
                        }));

                    page.Footer(f => f.AlignCenter().Text(t => t
                        .Text("OfficeIMO.Pdf compose - page ")
                        .CurrentPage()
                        .Text("/")
                        .TotalPages()));
                });
            });

            doc.Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
