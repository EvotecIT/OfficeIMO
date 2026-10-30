using OfficeIMO.Pdf;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class ShowcaseStatementPdf {
        public static void Example_Pdf_ShowcaseStatement(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.Showcase.Statement.pdf");

            var options = new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10.5),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(24, 31, 42),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = global::OfficeIMO.OfficeLength.Points(8),
                HeaderFormat = "OfficeIMO.Pdf showcase - business statement",
                HeaderAlign = PdfAlign.Left,
                HeaderOffsetY = global::OfficeIMO.OfficeLength.Points(18),
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = global::OfficeIMO.OfficeLength.Points(9),
                FooterFormat = "{page} / {pages}",
                FooterAlign = PdfAlign.Center,
                ShowPageNumbers = true,
                CreateOutlineFromHeadings = true
            };

            PdfDocument.Create(document => {
                    document.Page(page => {
                        page.Margin(global::OfficeIMO.OfficeLength.Points(54), global::OfficeIMO.OfficeLength.Points(54), global::OfficeIMO.OfficeLength.Points(54), global::OfficeIMO.OfficeLength.Points(58));
                        page.Content(content => {
                            content.Item(item => item.H1("Statement #4048", PdfAlign.Left, global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42)));
                            content.Item(item => item.Paragraph(p => p.Text("Issue date: 23/12/2025").Text("   Due date: 06/01/2026"), style: TightParagraph()));
                            content.Item(item => item.Shape(CreateBrandRule(), PdfAlign.Right, spacingBefore: global::OfficeIMO.OfficeLength.Points(6), spacingAfter: global::OfficeIMO.OfficeLength.Points(22)));
                            content.Row(row => {
                                row.Gap(global::OfficeIMO.OfficeLength.Points(58))
                                    .PercentColumn(50, column => {
                                        column.Paragraph(p => p.Bold("From"), style: LabelParagraph());
                                        column.HR(global::OfficeIMO.OfficeLength.Points(1.2), global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42), spacingBefore: global::OfficeIMO.OfficeLength.Points(0), spacingAfter: global::OfficeIMO.OfficeLength.Points(8));
                                        column.Paragraph(p => p.Text("Syllabae Representative\nOvum picem\nPrinceps avem distant, Linteum amicitia\nofficium21@aut statum.com\n881-306-3914"), style: AddressParagraph());
                                    })
                                    .PercentColumn(50, column => {
                                        column.Paragraph(p => p.Bold("For"), style: LabelParagraph());
                                        column.HR(global::OfficeIMO.OfficeLength.Points(1.2), global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42), spacingBefore: global::OfficeIMO.OfficeLength.Points(0), spacingAfter: global::OfficeIMO.OfficeLength.Points(8));
                                        column.Paragraph(p => p.Text("Ceciderit Original\nAurum currunt\nSolis multum platea, Cocus fuge fluvio\nsubsisto93@celeritate.com\n839-621-9110"), style: AddressParagraph());
                                    });
                            });

                            content.Spacer(global::OfficeIMO.OfficeLength.Points(24));
                            content.Item(item => item.Table(CreateLineItemRows(), style: LineItemStyle()));
                            content.Spacer(global::OfficeIMO.OfficeLength.Points(10));
                            content.Row(row => {
                                row.Style(new PdfRowStyle { KeepTogether = true, SpacingBefore = global::OfficeIMO.OfficeLength.Points(2) })
                                    .Gap(global::OfficeIMO.OfficeLength.Points(22))
                                    .PercentColumn(58, column => {
                                        column.PanelParagraph(
                                            p => p.Bold("Payment note: ").Text("This is a generic document sample. The layout is built from headings, paragraphs, rows, tables, and cell styles rather than an invoice-specific API."),
                                            new PdfPanelStyle {
                                                Background = global::OfficeIMO.OfficeColor.FromRgb(247, 250, 252),
                                                BorderColor = global::OfficeIMO.OfficeColor.FromRgb(211, 219, 229),
                                                BorderWidth = global::OfficeIMO.OfficeLength.Points(0.6),
                                                PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                                                PaddingY = global::OfficeIMO.OfficeLength.Points(8)
                                            },
                                            defaultColor: global::OfficeIMO.OfficeColor.FromRgb(51, 65, 85));
                                    })
                                    .PercentColumn(42, column => {
                                        column.Table(CreateTotalsRows(), style: TotalsStyle());
                                    });
                            });
                        });
                    });
                }, options)
                .Meta(
                    title: "OfficeIMO.Pdf Showcase Statement",
                    author: "OfficeIMO",
                    subject: "Generic business document built from primitive PDF blocks",
                    keywords: "OfficeIMO,Pdf,statement,tables,rows")
                .Save(path);

            if (open) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
            }
        }

        private static string FormatMoney(decimal value) {
            return value.ToString("0.00", CultureInfo.InvariantCulture).Replace('.', ',') + " PLN";
        }

        private static string FormatVat(decimal value) {
            return value.ToString("0.0", CultureInfo.InvariantCulture).Replace('.', ',') + "%";
        }

        private static PdfTableCell[][] CreateLineItemRows() {
            var names = new[] {
                "Experientiam nostrum",
                "Radio insulam pluviae",
                "Sufficit consilium",
                "Rex maxime Dixitque",
                "Vitulus exspecta",
                "Aliquis sentio",
                "Cum anno deserto",
                "Splendidum etiamne fac",
                "Stagnum fune",
                "Tabula ipse",
                "Actum exemplum princeps",
                "Dimidiam ipsum",
                "Coegi aliquid",
                "Pauper tenuis",
                "Bigas rotam dicunt",
                "Custodi puella",
                "Praestare eorum",
                "Umero certus tantum"
            };
            var prices = new[] { 31.80m, 62.57m, 42.50m, 22.75m, 85.56m, 40.58m, 37.72m, 88.21m, 59.83m, 21.85m, 6.41m, 93.57m, 77.27m, 9.68m, 23.94m, 79.05m, 11.65m, 81.72m };
            var quantities = new[] { 2, 7, 5, 5, 2, 7, 7, 5, 1, 9, 8, 7, 1, 8, 6, 8, 7, 2 };
            var rows = new PdfTableCell[names.Length + 1][];
            rows[0] = new[] {
                PdfTableCell.TextCell("#"),
                PdfTableCell.TextCell("Product"),
                PdfTableCell.TextCell("Unit price"),
                PdfTableCell.TextCell("Quantity"),
                PdfTableCell.TextCell("Total")
            };

            for (int i = 0; i < names.Length; i++) {
                decimal total = prices[i] * quantities[i];
                rows[i + 1] = new[] {
                    PdfTableCell.TextCell((i + 1).ToString(CultureInfo.InvariantCulture)),
                    PdfTableCell.TextCell(names[i]),
                    PdfTableCell.TextCell(FormatMoney(prices[i])),
                    PdfTableCell.TextCell(quantities[i].ToString(CultureInfo.InvariantCulture)),
                    PdfTableCell.TextCell(FormatMoney(total))
                };
            }

            return rows;
        }

        private static PdfTableCell[][] CreateTotalsRows() {
            const decimal subtotal = 4122.59m;
            const decimal vat = 948.20m;
            const decimal total = 5070.79m;

            return new[] {
                new[] { PdfTableCell.Span("Summary", 2) },
                new[] { PdfTableCell.TextCell("Subtotal"), PdfTableCell.TextCell(FormatMoney(subtotal)) },
                new[] { PdfTableCell.TextCell("VAT " + FormatVat(23m)), PdfTableCell.TextCell(FormatMoney(vat)) },
                new[] { PdfTableCell.TextCell("Total"), PdfTableCell.TextCell(FormatMoney(total)) }
            };
        }

        private static PdfTableStyle LineItemStyle() {
            return new PdfTableStyle {
                BorderColor = null,
                BorderWidth = global::OfficeIMO.OfficeLength.Points(0),
                HeaderFill = null,
                HeaderTextColor = global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42),
                HeaderSeparatorColor = global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42),
                HeaderSeparatorWidth = global::OfficeIMO.OfficeLength.Points(1.2),
                RowSeparatorColor = global::OfficeIMO.OfficeColor.FromRgb(219, 226, 235),
                RowSeparatorWidth = global::OfficeIMO.OfficeLength.Points(0.55),
                RowStripeFill = null,
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(2),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(5),
                HeaderFontSize = global::OfficeIMO.OfficeLength.Points(10),
                FontSize = global::OfficeIMO.OfficeLength.Points(10),
                LineHeight = 1.2,
                RightAlignNumeric = true,
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(26), null, global::OfficeIMO.OfficeLength.Points(92), global::OfficeIMO.OfficeLength.Points(62), global::OfficeIMO.OfficeLength.Points(92) },
                ColumnWidthWeights = new List<double> { 0.4, 4.2, 1.3, 0.9, 1.3 },
                Alignments = new List<PdfColumnAlign> {
                    PdfColumnAlign.Left,
                    PdfColumnAlign.Left,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Right
                }
            };
        }

        private static PdfTableStyle TotalsStyle() {
            return new PdfTableStyle {
                BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225),
                BorderWidth = global::OfficeIMO.OfficeLength.Points(0.4),
                HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42),
                HeaderTextColor = global::OfficeIMO.OfficeColor.White,
                HeaderRowCount = 1,
                RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(8),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(6),
                HeaderFontSize = global::OfficeIMO.OfficeLength.Points(10),
                FontSize = global::OfficeIMO.OfficeLength.Points(10),
                RightAlignNumeric = true,
                Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Right },
                ColumnWidthWeights = new List<double> { 1.2, 1.0 }
            };
        }

        private static PdfParagraphStyle TightParagraph() {
            return new PdfParagraphStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(2), LineHeight = 1.15 };
        }

        private static PdfParagraphStyle LabelParagraph() {
            return new PdfParagraphStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), LineHeight = 1.1 };
        }

        private static PdfParagraphStyle AddressParagraph() {
            return new PdfParagraphStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), LineHeight = 1.35 };
        }

        private static OfficeIMO.Drawing.OfficeShape CreateBrandRule() {
            var shape = OfficeIMO.Drawing.OfficeShape.RoundedRectangle(156, 8, 4);
            shape.FillGradient = OfficeIMO.Drawing.OfficeLinearGradient.Horizontal(
                OfficeIMO.OfficeColor.FromRgb(15, 23, 42),
                OfficeIMO.OfficeColor.FromRgb(14, 165, 233));
            shape.StrokeWidth = 0;
            return shape;
        }
    }
}
