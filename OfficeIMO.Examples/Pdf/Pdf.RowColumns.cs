using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class RowColumnsPdf {
        public static void Example_Pdf_RowColumns(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.RowColumns.pdf");

            var options = new PdfCreateOptions {
                HeaderFormat = "OfficeIMO.Pdf row columns",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            };

            PdfDocument.Create(document => {
                    document.Page(page => {
                        page.Content(content => {
                            content.Column(column => {
                                column.Item().H1("Row Columns");
                                column.Item().Paragraph(p => p.Text("A compact visual gate for explicit fixed, content-sized, and relative columns with first-class gutters."));
                            });
                            content.Row(row => {
                                row.Gap(18);
                                row.FixedColumn(112, column => column
                                    .H2("Status")
                                    .Paragraph(p => p
                                        .Text("The left column carries operational copy with comfortable wrapping, spacing, and no collision with the neighboring column."))
                                    .Bullets(new[] {
                                        "Gutters are first-class layout state.",
                                        "List markers stay inside the column frame."
                                    }, color: PdfColor.FromRgb(55, 65, 81))
                                    .PanelParagraph(
                                        p => p.Bold("Callout: ").Text("column panels can hold emphasis without leaving the row flow."),
                                        new PdfPanelStyle {
                                            Background = PdfColor.FromRgb(248, 250, 252),
                                            BorderColor = PdfColor.FromRgb(183, 194, 207),
                                            PaddingX = 7,
                                            PaddingY = 5,
                                            KeepTogether = true
                                        })
                                    .RoundedRectangle(96, 5, 2.5, strokeColor: PdfColor.FromRgb(22, 101, 52), strokeWidth: 0, fillColor: PdfColor.FromRgb(22, 163, 74), spacingBefore: 8, spacingAfter: 8)
                                    .Paragraph(p => p
                                        .Bold("Ready: ")
                                        .Text("row gutters are part of the composition model instead of caller-managed whitespace.")));
                                row.AutoColumn(column => column
                                    .H2("Owner")
                                    .Text("Operations"), maximum: 82);
                                row.RelativeColumn(column => column
                                    .H2("Next")
                                    .Paragraph(p => p
                                        .Text("The right column uses the same page flow but starts after an explicit gutter, giving report layouts a professional reading rhythm."))
                                    .Numbered(new[] {
                                        "Compose column content.",
                                        "Render each list item independently.",
                                        "Compare the raster baseline."
                                    }, color: PdfColor.FromRgb(55, 65, 81))
                                    .Table(new[] {
                                        new[] { "Metric", "Value" },
                                        new[] { "Sizing", "Relative" },
                                        new[] { "Panels", "Yes" }
                                    }, style: CreateTableStyle())
                                    .HR(0.8, PdfColor.FromRgb(183, 194, 207), 8, 8)
                                    .Paragraph(p => p
                                        .Bold("Review: ")
                                        .Text("columns keep long notes readable without placing them over neighboring content.")));
                            });
                            content.Column(column => {
                                column.Item().Paragraph(p => p.Text("End of row column sample."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80));
                            });
                        });
                    });
                }, options)
                .Meta(title: "OfficeIMO.Pdf Row Columns", author: "OfficeIMO")
                .Save(path);

            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }

        private static PdfTableStyle CreateTableStyle() {
            var style = PdfTheme.FromOfficeTheme(OfficeTheme.Default).TableStyle!;
            style.CellPaddingX = 4;
            style.CellPaddingY = 3;
            style.CellPaddingTop = 3;
            style.CellPaddingBottom = 3;
            style.HeaderRowCount = 1;
            style.RightAlignNumeric = false;
            style.SpacingBefore = 7;
            style.SpacingAfter = 7;
            style.Alignments = new System.Collections.Generic.List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Right };
            style.ColumnWidthWeights = new System.Collections.Generic.List<double> { 1.2, 0.8 };
            return style;
        }
    }
}
