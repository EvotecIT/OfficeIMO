using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class ShowcaseDashboardPdf {
        public static void Example_Pdf_ShowcaseDashboard(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.Showcase.Dashboard.pdf");

            var options = new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 9.5,
                DefaultTextColor = PdfColor.FromRgb(30, 41, 59),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "Northwind Support \u2013 Q3 operations review (sample data)",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "Page {page} of {pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true,
                CreateOutlineFromHeadings = true
            };

            PdfDocument.Create(document => {
                    #region Excerpt:executive-pdf-dashboard
                    document.Page(page => {
                        page.Size(PdfPageSizes.A4).Landscape().Margin(42, 38, 42, 42);
                        page.DefaultParagraphStyle(new PdfParagraphStyle { LineHeight = 1.18, SpacingAfter = 5 });
                        page.Content(content => {
                            content.Item(item => item.H1("Quarterly operations dashboard", PdfAlign.Left, PdfColor.FromRgb(15, 23, 42)));
                            content.Item(item => item.Paragraph(p => p.Text("Service levels, resolved volume and the decisions needed for Q4."), style: new PdfParagraphStyle { SpacingAfter = 10 }));

                            // Four KPI panels share the row width.
                            content.Row(row => {
                                row.Gap(14)
                                    .PercentColumn(25, column => column.PanelParagraph(p => p.FontSize(22).Bold("94.2%").FontSize(9.5).Text("\nSLA attainment (target 92%)"), MetricPanel(PdfColor.FromRgb(236, 253, 245), PdfColor.FromRgb(22, 163, 74)), PdfAlign.Left, PdfColor.FromRgb(22, 101, 52)))
                                    .PercentColumn(25, column => column.PanelParagraph(p => p.FontSize(22).Bold("1.8 h").FontSize(9.5).Text("\nMean first response"), MetricPanel(PdfColor.FromRgb(239, 246, 255), PdfColor.FromRgb(37, 99, 235)), PdfAlign.Left, PdfColor.FromRgb(30, 64, 175)))
                                    .PercentColumn(25, column => column.PanelParagraph(p => p.FontSize(22).Bold("37").FontSize(9.5).Text("\nOpen escalations"), MetricPanel(PdfColor.FromRgb(255, 251, 235), PdfColor.FromRgb(217, 119, 6)), PdfAlign.Left, PdfColor.FromRgb(146, 64, 14)))
                                    .PercentColumn(25, column => column.PanelParagraph(p => p.FontSize(22).Bold("0").FontSize(9.5).Text("\nCritical blockers"), MetricPanel(PdfColor.FromRgb(248, 250, 252), PdfColor.FromRgb(100, 116, 139)), PdfAlign.Left, PdfColor.FromRgb(51, 65, 85)));
                            });

                            content.Spacer(12);

                            // Chart and table on the left, narrative and decisions on the right.
                            content.Row(row => {
                                row.Gap(18)
                                    .PercentColumn(58, column => {
                                        column.Paragraph(p => p.Bold("Tickets resolved per month"), style: SectionLabel());
                                        column.Drawing(CreateTrendDrawing(), PdfAlign.Left, spacingBefore: 2, spacingAfter: 8);
                                        column.Table(CreateServiceRows(), style: ServiceTableStyle());
                                    })
                                    .PercentColumn(42, column => {
                                        column.Paragraph(p => p.Bold("Summary"), style: SectionLabel());
                                        column.PanelParagraph(
                                            p => p.Text("SLA attainment finished above the 92% target in all three months while resolved volume grew 15%. Integrations is the only service line below target, and its response time is the main driver."),
                                            new PdfPanelStyle {
                                                Background = PdfColor.FromRgb(248, 250, 252),
                                                BorderColor = PdfColor.FromRgb(203, 213, 225),
                                                BorderWidth = 0.7,
                                                PaddingX = 10,
                                                PaddingY = 8
                                            });
                                        column.Bullets(new[] {
                                            "Resolved tickets: 2,617, up from 2,274 last quarter.",
                                            "Security and Billing answer in about an hour.",
                                            "Integrations answers in 3.1 h against a 2 h goal.",
                                            "No critical blockers are open at quarter end."
                                        }, style: new PdfListStyle { SpacingAfter = 4, ItemSpacing = 2 });
                                        column.Table(CreateDecisionRows(), style: DecisionTableStyle());
                                    });
                            });
                        });
                    });
                    #endregion
                }, options)
                .Meta(
                    title: "Quarterly operations dashboard",
                    author: "Northwind Support",
                    subject: "Q3 operations review (sample data)",
                    keywords: "dashboard,operations,SLA,quarterly")
                .Save(path);

            if (open) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
            }
        }

        private static PdfPanelStyle MetricPanel(PdfColor background, PdfColor border) {
            return new PdfPanelStyle {
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

        // The drawing spans the left column, so its right edge lines up with the table below it.
        private const double ChartWidth = 429;

        private static OfficeDrawing CreateTrendDrawing() {
            const double height = 176, plotLeft = 40, plotRight = ChartWidth - 12, baseline = 148, ticketsPerPoint = 8;
            OfficeThemePalette palette = OfficeTheme.Default.Palette;
            var drawing = new OfficeDrawing(ChartWidth, height);
            var label = new OfficeFontInfo("Helvetica", 8);
            OfficeColor muted = palette.TextMuted;
            OfficeColor text = palette.Text;

            var frame = OfficeShape.Rectangle(ChartWidth, height);
            frame.FillColor = palette.Background;
            frame.StrokeColor = OfficeColor.FromRgb(220, 226, 234);
            frame.StrokeWidth = 0.75;
            drawing.AddShape(frame, 0, 0);

            // Light gridlines with their values on the axis; the baseline is the zero line.
            for (int value = 0; value <= 1000; value += 250) {
                double y = baseline - value / ticketsPerPoint;
                var grid = OfficeShape.Line(0, 0, plotRight - plotLeft, 0);
                grid.StrokeColor = value == 0 ? OfficeColor.FromRgb(201, 208, 218) : OfficeColor.FromRgb(230, 234, 240);
                grid.StrokeWidth = 0.5;
                drawing.AddShape(grid, plotLeft, y);
                drawing.AddText(value.ToString("N0", CultureInfo.InvariantCulture), 4, y - 5, plotLeft - 10, 10, label, muted, OfficeTextAlignment.Right);
            }

            string[] months = { "July", "August", "September" };
            int[] resolved = { 812, 871, 934 };
            double slot = (plotRight - plotLeft) / resolved.Length;
            for (int i = 0; i < resolved.Length; i++) {
                double barHeight = resolved[i] / ticketsPerPoint;
                double x = plotLeft + i * slot + slot * 0.25;
                var bar = OfficeShape.Rectangle(slot * 0.5, barHeight);
                bar.FillColor = palette.Accent1;
                bar.StrokeWidth = 0;
                drawing.AddShape(bar, x, baseline - barHeight);
                // Data labels sit inside the top of each bar; white on the accent fill passes WCAG AA,
                // and the target line above the shorter bars stays clear of them.
                drawing.AddText(resolved[i].ToString("N0", CultureInfo.InvariantCulture), x, baseline - barHeight + 5, slot * 0.5, 11, new OfficeFontInfo("Helvetica", 9, OfficeFontStyle.Bold), palette.Background, OfficeTextAlignment.Center);
                drawing.AddText(months[i], x, baseline + 6, slot * 0.5, 11, label, muted, OfficeTextAlignment.Center);
            }

            // Target line: 850 tickets a month, labelled at its left end clear of the bars.
            double targetY = baseline - 850 / ticketsPerPoint;
            var target = OfficeShape.Line(0, 0, plotRight - plotLeft, 0);
            target.StrokeColor = text;
            target.StrokeWidth = 1;
            target.StrokeDashStyle = OfficeStrokeDashStyle.Dash;
            drawing.AddShape(target, plotLeft, targetY);
            drawing.AddText("Target 850", plotLeft + 4, targetY - 22, 70, 10, label, text, OfficeTextAlignment.Left);

            return drawing;
        }

        private static PdfTableCell[][] CreateServiceRows() {
            OfficeThemePalette palette = OfficeTheme.Default.Palette;
            PdfTableCell Status(string status) => PdfTableCell.RichTextCell(new[] {
                PdfTextRun.Bolded(status, PdfColor.FromOfficeColor(status == "Watch" ? palette.Warning : palette.Success))
            });
            PdfTableCell T(string value) => PdfTableCell.TextCell(value);
            return new[] {
                new[] { T("Service line"), T("SLA"), T("Response"), T("Change"), T("Status") },
                new[] { T("Billing"), T("96.1%"), T("1.2 h"), T("+2.4 pts"), Status("On target") },
                new[] { T("Onboarding"), T("93.4%"), T("1.9 h"), T("+1.1 pts"), Status("On target") },
                new[] { T("Integrations"), T("89.7%"), T("3.1 h"), T("-1.8 pts"), Status("Watch") },
                new[] { T("Security"), T("97.0%"), T("0.9 h"), T("+0.3 pts"), Status("On target") }
            };
        }
        private static PdfTableCell[][] CreateDecisionRows() {
            return new[] {
                new[] { PdfTableCell.Span("Decisions for Q4", 2) },
                new[] { PdfTableCell.TextCell("Staffing"), PdfTableCell.TextCell("Move two engineers to Integrations") },
                new[] { PdfTableCell.TextCell("Escalations"), PdfTableCell.TextCell("Review the 37 open items every week") },
                new[] { PdfTableCell.TextCell("Targets"), PdfTableCell.TextCell("Keep 92% SLA; tighten response goal to 1.5 h") }
            };
        }

        private static PdfTableStyle ServiceTableStyle() {
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
                ColumnWidthWeights = new List<double> { 1.6, 0.8, 1.0, 1.0, 1.1 },
                Alignments = new List<PdfColumnAlign> {
                    PdfColumnAlign.Left,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Right,
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
