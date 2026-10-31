using OfficeIMO.Pdf;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class ShowcaseStatementPdf {
        public static void Example_Pdf_ShowcaseStatement(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.Showcase.Statement.pdf");

            var options = new PdfCreateOptions {
                HeaderFormat = "OfficeIMO.Pdf showcase - business statement",
                HeaderAlign = PdfAlign.Left,
                HeaderOffsetY = global::OfficeIMO.OfficeLength.Points(18),
                ShowHeader = true,
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true,
                CreateOutlineFromHeadings = true
            };

            PdfDocument.Create(document => {
                    document.Page(page => {
                        page.Margin(global::OfficeIMO.OfficeLength.Points(54), global::OfficeIMO.OfficeLength.Points(54), global::OfficeIMO.OfficeLength.Points(54), global::OfficeIMO.OfficeLength.Points(58));
                        page.Content(content => {
                            content.Item(item => item.H1("Statement #4048", PdfAlign.Left));
                            content.Item(item => item.Paragraph(p => p.Text("Issue date: 23/12/2025 \u00B7 Due date: 06/01/2026"), style: TightParagraph()));
                            content.Spacer(OfficeLength.Points(18));
                            content.Row(row => {
                                row.Gap(global::OfficeIMO.OfficeLength.Points(58))
                                    .PercentColumn(50, column => {
                                        column.Paragraph(p => p.Bold("From"), style: LabelParagraph());
                                        column.HR(OfficeLength.Points(1.2), OfficeColor.FromRgb(15, 23, 42), spacingBefore: OfficeLength.Points(0), spacingAfter: OfficeLength.Points(8));
                                        column.Paragraph(p => p.Text("Syllabae Representative\nOvum picem\nPrinceps avem distant, Linteum amicitia\nofficium21@autstatum.com\n881-306-3914"), style: AddressParagraph());
                                    })
                                    .PercentColumn(50, column => {
                                        column.Paragraph(p => p.Bold("For"), style: LabelParagraph());
                                        column.HR(global::OfficeIMO.OfficeLength.Points(1.2), global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42), spacingBefore: global::OfficeIMO.OfficeLength.Points(0), spacingAfter: global::OfficeIMO.OfficeLength.Points(8));
                                        column.Paragraph(p => p.Text("Ceciderit Original\nAurum currunt\nSolis multum platea, Cocus fuge fluvio\nsubsisto93@celeritate.com\n839-621-9110"), style: AddressParagraph());
                                    });
                            });

                            content.Spacer(OfficeLength.Points(24));
                            content.Table(CreateLineItemRows(), style: LineItemStyle());
                            content.Row(row => {
                                row.Style(new PdfRowStyle { KeepTogether = true, SpacingBefore = OfficeLength.Points(10) })
                                    .Gap(OfficeLength.Points(22))
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
            var style = PdfTheme.FromOfficeTheme(OfficeTheme.Default).TableStyle!;
            style.HeaderFill = null;
            style.HeaderSeparatorColor = OfficeColor.FromRgb(15, 23, 42);
            style.HeaderSeparatorWidth = OfficeLength.Points(1.2);
            style.RowStripeFill = null;
            style.CellPaddingX = OfficeLength.Points(6);
            style.CellPaddingY = OfficeLength.Points(5);
            style.CellPaddingTop = OfficeLength.Points(5);
            style.CellPaddingBottom = OfficeLength.Points(5);
            style.FontSize = OfficeLength.Points(10);
            style.LineHeight = 1.2;
            style.RightAlignNumeric = true;
            // The totals beside the payment note follow the final rows instead of starting a page alone.
            style.KeepWithNext = true;
            style.ColumnWidths = new List<OfficeLength?> { OfficeLength.Points(26), null, OfficeLength.Points(92), OfficeLength.Points(62), OfficeLength.Points(92) };
            style.ColumnWidthWeights = new List<double> { 0.4, 4.2, 1.3, 0.9, 1.3 };
            style.Alignments = new List<PdfColumnAlign> {
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Left,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Right
                };
            return style;
        }

        private static PdfTableStyle TotalsStyle() {
            var style = PdfTheme.FromOfficeTheme(OfficeTheme.Default).TableStyle!;
            style.HeaderRowCount = 1;
            style.FooterRowCount = 1;
            style.FooterBold = true;
            style.FooterFill = null;
            style.FooterSeparatorColor = OfficeColor.FromRgb(15, 23, 42);
            style.FooterSeparatorWidth = OfficeLength.Points(1);
            style.CellPaddingX = OfficeLength.Points(8);
            style.CellPaddingY = OfficeLength.Points(6);
            style.CellPaddingTop = OfficeLength.Points(6);
            style.CellPaddingBottom = OfficeLength.Points(6);
            style.FontSize = OfficeLength.Points(10);
            style.RightAlignNumeric = true;
            style.Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Right };
            style.ColumnWidthWeights = new List<double> { 1.2, 1.0 };
            return style;
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
    }
}
