using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class FlowDslDemo {
        public static void Example_Pdf_FlowDslDemo(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.FlowDslDemo.pdf");

            var doc = PdfDoc.Create().Compose(document => {
                document.Page(page => {
                    page.Size(PageSizes.A5);
                    page.DefaultTextStyle(x => x.FontSize(20));
                    page.Margin(25);

                    page.Content(c => c
                        .PaddingBottom(15)
                        .Column(column => {
                            column.Item().H1("Report Title");
                            column.Item().PageBreak();
                            column.Item().Paragraph(p => p.Color(PdfColor.FromRgb(200,0,0)).Text("Red section paragraph."));
                            column.Item().PageBreak();
                            column.Item().Paragraph(p => p.Color(PdfColor.FromRgb(0,160,0)).Text("Green section paragraph."));
                            column.Item().PageBreak();
                            column.Item().Paragraph(p => p.Color(PdfColor.FromRgb(20,90,180)).Text("Blue section paragraph."));
                        }));

                    page.Footer(f => f.AlignCenter().PageNumber());
                });
            });

            doc.Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
