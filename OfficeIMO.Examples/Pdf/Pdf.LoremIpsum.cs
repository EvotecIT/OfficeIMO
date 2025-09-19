using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class LoremIpsumDemo {
        public static void Example_Pdf_LoremIpsum(string folderPath, bool open = false) {
            string path1 = Path.Combine(folderPath, "Pdf.Lorem.Plain.pdf");
            string path2 = Path.Combine(folderPath, "Pdf.Lorem.Panel.pdf");
            string path3 = Path.Combine(folderPath, "Pdf.Lorem.TwoColumns.pdf");

            string lipsum = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                            "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. " +
                            "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. " +
                            "Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum. ";
            // make it long
            string longText = string.Concat(lipsum, lipsum, lipsum, lipsum, lipsum);

            // 1) Plain paragraph
            PdfDoc.Create()
                .H1("Lorem Ipsum — Plain Paragraph", PdfAlign.Center)
                .Paragraph(p => p.Text(longText))
                .Save(path1);

            // 2) Inside panel
            PdfDoc.Create()
                .H1("Lorem Ipsum — Panel Paragraph", PdfAlign.Center)
                .PanelParagraph(p => p.Text(longText), new PanelStyle {
                    Background = new PdfColor(0.96, 0.96, 0.98),
                    BorderColor = PdfColor.FromRgb(210, 210, 210),
                    BorderWidth = 0.5,
                    PaddingX = 8,
                    PaddingY = 6,
                    MaxWidth = 380,
                    Align = PdfAlign.Center
                })
                .Save(path2);

            // 3) Two columns via Compose Row
            PdfDoc.Create().Compose(d => {
                d.Page(p => {
                    p.Size(PageSizes.Letter);
                    p.Margin(36);
                    p.DefaultTextStyle(s => s.FontSize(11));
                    p.Content(c => c
                        .Column(col => col.Item().H1("Lorem Ipsum — Two Columns"))
                        .Row(row => {
                            row.Column(50, col => col.Paragraph(pr => pr.Text(longText)));
                            row.Column(50, col => col.Paragraph(pr => pr.Text(longText)));
                        })
                    );
                });
            }).Save(path3);

            if (open) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path1, UseShellExecute = true });
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path2, UseShellExecute = true });
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path3, UseShellExecute = true });
            }
        }
    }
}

