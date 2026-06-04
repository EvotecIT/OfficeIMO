using OfficeIMO.Pdf;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class ShowcaseStatementPdf {
        public static void Example_Pdf_ShowcaseStatement(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.Showcase.Statement.pdf");

            PdfDocument.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10.5,
                    DefaultTextColor = PdfColor.FromRgb(24, 31, 42),
                    HeaderFont = PdfStandardFont.Helvetica,
                    HeaderFontSize = 8,
                    HeaderFormat = "OfficeIMO.Pdf showcase - business statement",
                    HeaderAlign = PdfAlign.Left,
                    HeaderOffsetY = 18,
                    ShowHeader = true,
                    FooterFont = PdfStandardFont.Helvetica,
                    FooterFontSize = 9,
                    FooterFormat = "{page} / {pages}",
                    FooterAlign = PdfAlign.Center,
                    ShowPageNumbers = true,
                    CreateOutlineFromHeadings = true
                })
                .Meta(
                    title: "OfficeIMO.Pdf Showcase Statement",
                    author: "OfficeIMO",
                    subject: "Generic business document built from primitive PDF blocks",
                    keywords: "OfficeIMO,Pdf,statement,tables,rows")
                .Compose(document => {
                    document.Page(page => {
                        page.Margin(54, 54, 54, 58);
                        page.Content(content => {
                            content.Item(item => item.H1("Statement #4048", PdfAlign.Left, PdfColor.FromRgb(15, 23, 42)));
                            content.Item(item => item.Paragraph(p => p.Text("Issue date: 23/12/2025").Text("   Due date: 06/01/2026"), style: TightParagraph()));
                            content.Item(item => item.Shape(CreateBrandRule(), PdfAlign.Right, spacingBefore: 6, spacingAfter: 22));
                            content.Row(row => {
                                row.Gap(58)
                                    .Column(50, column => {
                                        column.Paragraph(p => p.Bold("From"), style: LabelParagraph());
                                        column.HR(1.2, PdfColor.FromRgb(15, 23, 42), spacingBefore: 0, spacingAfter: 8);
                                        column.Paragraph(p => p.Text("Syllabae Representative\nOvum picem\nPrinceps avem distant, Linteum amicitia\nofficium21@aut statum.com\n881-306-3914"), style: AddressParagraph());
                                    })
                                    .Column(50, column => {
                                        column.Paragraph(p => p.Bold("For"), style: LabelParagraph());
                                        column.HR(1.2, PdfColor.FromRgb(15, 23, 42), spacingBefore: 0, spacingAfter: 8);
                                        column.Paragraph(p => p.Text("Ceciderit Original\nAurum currunt\nSolis multum platea, Cocus fuge fluvio\nsubsisto93@celeritate.com\n839-621-9110"), style: AddressParagraph());
                                    });
                            });

                            content.Spacer(24);
                            content.Item(item => item.Table(CreateLineItemRows(), style: LineItemStyle()));
                            content.Spacer(10);
                            content.Row(row => {
                                row.Style(new PdfRowStyle { KeepTogether = true, SpacingBefore = 2 })
                                    .Gap(22)
                                    .Column(58, column => {
                                        column.PanelParagraph(
                                            p => p.Bold("Payment note: ").Text("This is a generic document sample. The layout is built from headings, paragraphs, rows, tables, and cell styles rather than an invoice-specific API."),
                                            new PanelStyle {
                                                Background = PdfColor.FromRgb(247, 250, 252),
                                                BorderColor = PdfColor.FromRgb(211, 219, 229),
                                                BorderWidth = 0.6,
                                                PaddingX = 9,
                                                PaddingY = 8
                                            },
                                            defaultColor: PdfColor.FromRgb(51, 65, 85));
                                    })
                                    .Column(42, column => {
                                        column.Table(CreateTotalsRows(), style: TotalsStyle());
                                    });
                            });
                        });
                    });
                })
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
                BorderWidth = 0,
                HeaderFill = null,
                HeaderTextColor = PdfColor.FromRgb(15, 23, 42),
                HeaderSeparatorColor = PdfColor.FromRgb(15, 23, 42),
                HeaderSeparatorWidth = 1.2,
                RowSeparatorColor = PdfColor.FromRgb(219, 226, 235),
                RowSeparatorWidth = 0.55,
                RowStripeFill = null,
                CellPaddingX = 2,
                CellPaddingY = 5,
                HeaderFontSize = 10,
                FontSize = 10,
                LineHeight = 1.2,
                RightAlignNumeric = true,
                ColumnWidthPoints = new List<double?> { 26, null, 92, 62, 92 },
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
                BorderColor = PdfColor.FromRgb(203, 213, 225),
                BorderWidth = 0.4,
                HeaderFill = PdfColor.FromRgb(15, 23, 42),
                HeaderTextColor = PdfColor.White,
                HeaderRowCount = 1,
                RowStripeFill = PdfColor.FromRgb(248, 250, 252),
                CellPaddingX = 8,
                CellPaddingY = 6,
                HeaderFontSize = 10,
                FontSize = 10,
                RightAlignNumeric = true,
                Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Right },
                ColumnWidthWeights = new List<double> { 1.2, 1.0 }
            };
        }

        private static PdfParagraphStyle TightParagraph() {
            return new PdfParagraphStyle { SpacingBefore = 0, SpacingAfter = 2, LineHeight = 1.15 };
        }

        private static PdfParagraphStyle LabelParagraph() {
            return new PdfParagraphStyle { SpacingBefore = 0, SpacingAfter = 0, LineHeight = 1.1 };
        }

        private static PdfParagraphStyle AddressParagraph() {
            return new PdfParagraphStyle { SpacingBefore = 0, SpacingAfter = 0, LineHeight = 1.35 };
        }

        private static OfficeIMO.Drawing.OfficeShape CreateBrandRule() {
            var shape = OfficeIMO.Drawing.OfficeShape.RoundedRectangle(156, 8, 4);
            shape.FillGradient = OfficeIMO.Drawing.OfficeLinearGradient.Horizontal(
                OfficeIMO.Drawing.OfficeColor.FromRgb(15, 23, 42),
                OfficeIMO.Drawing.OfficeColor.FromRgb(14, 165, 233));
            shape.StrokeWidth = 0;
            return shape;
        }
    }
}
