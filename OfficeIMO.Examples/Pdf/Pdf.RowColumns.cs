using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class RowColumnsPdf {
        public static void Example_Pdf_RowColumns(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.RowColumns.pdf");

            PdfDocument.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10,
                    DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                    HeaderFont = PdfStandardFont.Helvetica,
                    HeaderFontSize = 8,
                    HeaderFormat = "OfficeIMO.Pdf row columns",
                    HeaderAlign = PdfAlign.Left,
                    ShowHeader = true,
                    FooterFont = PdfStandardFont.Helvetica,
                    FooterFontSize = 8,
                    FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                    FooterAlign = PdfAlign.Right,
                    ShowPageNumbers = true
                })
                .Meta(title: "OfficeIMO.Pdf Row Columns", author: "OfficeIMO")
                .Compose(document => {
                    document.Page(page => {
                        page.Content(content => {
                            content.Column(column => {
                                column.Item().H1("Row Columns");
                                column.Item().Paragraph(p => p.Text("A compact visual gate for composed columns with first-class gutters and independent column flow."));
                            });
                            content.Row(row => {
                                row.Gap(18);
                                row.Column(50, column => column
                                    .H2("Status")
                                    .Paragraph(p => p
                                        .Text("The left column carries operational copy with comfortable wrapping, spacing, and no collision with the neighboring column."))
                                    .Bullets(new[] {
                                        "Gutters are first-class layout state.",
                                        "List markers stay inside the column frame."
                                    }, color: PdfColor.FromRgb(55, 65, 81))
                                    .PanelParagraph(
                                        p => p.Bold("Callout: ").Text("column panels can hold emphasis without leaving the row flow."),
                                        new PanelStyle {
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
                                row.Column(50, column => column
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
                                        new[] { "Gutter", "18 pt" },
                                        new[] { "Panels", "Yes" }
                                    }, style: new PdfTableStyle {
                                        HeaderFill = PdfColor.FromRgb(25, 55, 85),
                                        HeaderTextColor = PdfColor.White,
                                        RowStripeFill = PdfColor.FromRgb(248, 250, 252),
                                        BorderColor = PdfColor.FromRgb(183, 194, 207),
                                        BorderWidth = 0.5,
                                        CellPaddingX = 4,
                                        CellPaddingY = 3,
                                        HeaderRowCount = 1,
                                        RightAlignNumeric = false,
                                        SpacingBefore = 7,
                                        SpacingAfter = 7,
                                        Alignments = new System.Collections.Generic.List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Right },
                                        ColumnWidthWeights = new System.Collections.Generic.List<double> { 1.2, 0.8 }
                                    })
                                    .HR(0.8, PdfColor.FromRgb(183, 194, 207), 8, 8)
                                    .Paragraph(p => p
                                        .Bold("Guarded: ")
                                        .Text("the Poppler baseline catches cramped columns and accidental gutter regressions.")));
                            });
                            content.Column(column => {
                                column.Item().Paragraph(p => p.Text("End of row column sample."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80));
                            });
                        });
                    });
                })
                .Save(path);

            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
