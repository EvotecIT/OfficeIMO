using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class ShowcaseDashboardPdf {
        public static void Example_Pdf_ShowcaseDashboard(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.Showcase.Dashboard.pdf");

            PdfDocument.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 9.5,
                    DefaultTextColor = PdfColor.FromRgb(30, 41, 59),
                    HeaderFont = PdfStandardFont.Helvetica,
                    HeaderFontSize = 8,
                    HeaderFormat = "OfficeIMO.Pdf showcase - landscape dashboard",
                    HeaderAlign = PdfAlign.Left,
                    ShowHeader = true,
                    FooterFont = PdfStandardFont.Helvetica,
                    FooterFontSize = 8,
                    FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                    FooterAlign = PdfAlign.Right,
                    ShowPageNumbers = true,
                    CreateOutlineFromHeadings = true
                })
                .Meta(
                    title: "OfficeIMO.Pdf Showcase Dashboard",
                    author: "OfficeIMO",
                    subject: "Landscape dashboard from generic PDF primitives",
                    keywords: "OfficeIMO,Pdf,dashboard,drawings,tables")
                .Compose(document => {
                    document.Page(page => {
                        page.Size(PageSizes.A4).Landscape().Margin(42, 38, 42, 42);
                        page.DefaultParagraphStyle(new PdfParagraphStyle { LineHeight = 1.18, SpacingAfter = 5 });
                        page.Content(content => {
                            content.Item(item => item.H1("Quarterly Operations Dashboard", PdfAlign.Left, PdfColor.FromRgb(15, 23, 42)));
                            content.Item(item => item.Paragraph(p => p.Text("A single-page control surface composed from rows, panels, reusable drawings, wrapped tables, and compact list rhythm."), style: new PdfParagraphStyle { SpacingAfter = 10 }));

                            content.Row(row => {
                                row.Gap(14)
                                    .Column(25, column => column.PanelParagraph(p => p.Bold("92%").Text("\nSLA attainment"), MetricPanel(PdfColor.FromRgb(236, 253, 245), PdfColor.FromRgb(22, 163, 74)), PdfAlign.Left, PdfColor.FromRgb(22, 101, 52)))
                                    .Column(25, column => column.PanelParagraph(p => p.Bold("1.8h").Text("\nMean response"), MetricPanel(PdfColor.FromRgb(239, 246, 255), PdfColor.FromRgb(37, 99, 235)), PdfAlign.Left, PdfColor.FromRgb(30, 64, 175)))
                                    .Column(25, column => column.PanelParagraph(p => p.Bold("34").Text("\nOpen actions"), MetricPanel(PdfColor.FromRgb(255, 251, 235), PdfColor.FromRgb(217, 119, 6)), PdfAlign.Left, PdfColor.FromRgb(146, 64, 14)))
                                    .Column(25, column => column.PanelParagraph(p => p.Bold("0").Text("\nCritical blockers"), MetricPanel(PdfColor.FromRgb(248, 250, 252), PdfColor.FromRgb(100, 116, 139)), PdfAlign.Left, PdfColor.FromRgb(51, 65, 85)));
                            });

                            content.Spacer(12);
                            content.Row(row => {
                                row.Gap(18)
                                    .Column(58, column => {
                                        column.Paragraph(p => p.Bold("Delivery trend"), style: SectionLabel());
                                        column.Drawing(CreateTrendDrawing(), PdfAlign.Left, spacingBefore: 2, spacingAfter: 8);
                                        column.Table(CreateRiskRows(), style: RiskTableStyle());
                                    })
                                    .Column(42, column => {
                                        column.Paragraph(p => p.Bold("Narrative"), style: SectionLabel());
                                        column.PanelParagraph(
                                            p => p.Text("The dashboard deliberately avoids a domain-specific report object. It uses the same primitive surface that a Word, Excel, or PowerPoint exporter could target later: page setup, rows, tables, paragraphs, shapes, and themes."),
                                            new PanelStyle {
                                                Background = PdfColor.FromRgb(248, 250, 252),
                                                BorderColor = PdfColor.FromRgb(203, 213, 225),
                                                BorderWidth = 0.7,
                                                PaddingX = 10,
                                                PaddingY = 8
                                            });
                                        column.Bullets(new[] {
                                            "Rows keep gutters as layout state.",
                                            "Tables use explicit widths and numeric alignment.",
                                            "Vector drawing comes from OfficeIMO.Drawing descriptors.",
                                            "Visual gates can rasterize the result and catch rhythm regressions."
                                        }, style: new PdfListStyle { SpacingAfter = 4, ItemSpacing = 2 });
                                        column.Table(CreateDecisionRows(), style: DecisionTableStyle());
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

        private static PanelStyle MetricPanel(PdfColor background, PdfColor border) {
            return new PanelStyle {
                Background = background,
                BorderColor = border,
                BorderWidth = 0.8,
                PaddingX = 10,
                PaddingY = 8
            };
        }

        private static PdfParagraphStyle SectionLabel() {
            return new PdfParagraphStyle { SpacingBefore = 0, SpacingAfter = 4, LineHeight = 1.1 };
        }

        private static OfficeDrawing CreateTrendDrawing() {
            var drawing = new OfficeDrawing(390, 168);

            var frame = OfficeShape.RoundedRectangle(390, 168, 8);
            frame.FillColor = OfficeColor.FromRgb(255, 255, 255);
            frame.StrokeColor = OfficeColor.FromRgb(203, 213, 225);
            frame.StrokeWidth = 0.8;
            drawing.AddShape(frame, 0, 0);

            for (int i = 0; i < 4; i++) {
                var grid = OfficeShape.Line(0, 0, 340, 0);
                grid.StrokeColor = OfficeColor.FromRgb(226, 232, 240);
                grid.StrokeWidth = 0.5;
                drawing.AddShape(grid, 28, 32 + i * 28);
            }

            double[] bars = { 72, 88, 58, 96, 110, 82 };
            for (int i = 0; i < bars.Length; i++) {
                var bar = OfficeShape.RoundedRectangle(30, bars[i], 4);
                bar.FillGradient = OfficeLinearGradient.Vertical(OfficeColor.FromRgb(14, 165, 233), OfficeColor.FromRgb(37, 99, 235));
                bar.StrokeWidth = 0;
                drawing.AddShape(bar, 44 + i * 46, 140 - bars[i]);
            }

            var target = OfficeShape.Line(0, 0, 306, 0);
            target.StrokeColor = OfficeColor.FromRgb(15, 23, 42);
            target.StrokeWidth = 1.2;
            target.StrokeDashStyle = OfficeStrokeDashStyle.Dash;
            drawing.AddShape(target, 44, 58);

            var marker = OfficeShape.Ellipse(18, 18);
            marker.FillColor = OfficeColor.FromRgb(220, 252, 231);
            marker.StrokeColor = OfficeColor.FromRgb(22, 163, 74);
            marker.StrokeWidth = 1.2;
            drawing.AddShape(marker, 326, 40);

            return drawing;
        }

        private static string[][] CreateRiskRows() {
            return new[] {
                new[] { "Area", "State", "Trend", "Owner" },
                new[] { "PDF layout rhythm", "Good", "+12%", "OfficeIMO.Pdf" },
                new[] { "Table wrapping", "Watch", "-3%", "Renderer" },
                new[] { "Read/manipulation", "Growing", "+31%", "Core" },
                new[] { "Word/Excel export path", "Planned", "+0%", "Roadmap" }
            };
        }

        private static PdfTableCell[][] CreateDecisionRows() {
            return new[] {
                new[] { PdfTableCell.Span("Next decisions", 2) },
                new[] { PdfTableCell.TextCell("Visual fixtures"), PdfTableCell.TextCell("Keep generic, use documents as gates") },
                new[] { PdfTableCell.TextCell("AST model"), PdfTableCell.TextCell("Promote page/content tree over helper-only APIs") },
                new[] { PdfTableCell.TextCell("Conversion"), PdfTableCell.TextCell("Add Word/Excel/PPT exporters in slices") }
            };
        }

        private static PdfTableStyle RiskTableStyle() {
            return new PdfTableStyle {
                HeaderFill = PdfColor.FromRgb(15, 23, 42),
                HeaderTextColor = PdfColor.White,
                RowStripeFill = PdfColor.FromRgb(248, 250, 252),
                BorderColor = PdfColor.FromRgb(203, 213, 225),
                BorderWidth = 0.4,
                RowSeparatorColor = PdfColor.FromRgb(226, 232, 240),
                RowSeparatorWidth = 0.45,
                CellPaddingX = 7,
                CellPaddingY = 5,
                HeaderFontSize = 9.5,
                FontSize = 9,
                RightAlignNumeric = true,
                ColumnWidthWeights = new List<double> { 2.1, 1.0, 0.8, 1.2 },
                Alignments = new List<PdfColumnAlign> {
                    PdfColumnAlign.Left,
                    PdfColumnAlign.Center,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Left
                }
            };
        }

        private static PdfTableStyle DecisionTableStyle() {
            return new PdfTableStyle {
                HeaderFill = PdfColor.FromRgb(30, 64, 175),
                HeaderTextColor = PdfColor.White,
                RowStripeFill = PdfColor.FromRgb(239, 246, 255),
                BorderColor = PdfColor.FromRgb(191, 219, 254),
                BorderWidth = 0.45,
                CellPaddingX = 7,
                CellPaddingY = 5,
                HeaderRowCount = 1,
                FontSize = 8.7,
                LineHeight = 1.15,
                ColumnWidthWeights = new List<double> { 1.0, 2.1 }
            };
        }
    }
}
