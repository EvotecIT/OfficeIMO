using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class ShowcaseDashboardPdf {
        public static void Example_Pdf_ShowcaseDashboard(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.Showcase.Dashboard.pdf");

            var options = new PdfCreateOptions {
                HeaderFormat = "Northwind Support \u2013 Q3 operations review (sample data)",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFormat = "Page {page} of {pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true,
                CreateOutlineFromHeadings = true
            };

            PdfDocument.Create(document => {
                    #region Excerpt:executive-pdf-dashboard
                    document.Page(page => {
                        page.Size(PdfPageSizes.A4).Landscape().Margin(global::OfficeIMO.OfficeLength.Points(42), global::OfficeIMO.OfficeLength.Points(38), global::OfficeIMO.OfficeLength.Points(42), global::OfficeIMO.OfficeLength.Points(42));
                        page.DefaultParagraphStyle(new PdfParagraphStyle { LineHeight = 1.18, SpacingAfter = global::OfficeIMO.OfficeLength.Points(5) });
                        page.Content(content => {
                            content.Item(item => item.H1("Quarterly operations dashboard", PdfAlign.Left));
                            content.Item(item => item.Paragraph(p => p.Text("Service levels, resolved volume and the decisions needed for Q4."), style: new PdfParagraphStyle { SpacingAfter = OfficeLength.Points(10) }));

                            // Four KPI panels share the row width.
                            content.Row(row => {
                                row.Gap(global::OfficeIMO.OfficeLength.Points(14))
                                    .PercentColumn(25, column => column.PanelParagraph(p => p.FontSize(global::OfficeIMO.OfficeLength.Points(22)).Bold("94.2%").FontSize(global::OfficeIMO.OfficeLength.Points(9.5)).Text("\nSLA attainment (target 92%)"), MetricPanel(global::OfficeIMO.OfficeColor.FromRgb(236, 253, 245), global::OfficeIMO.OfficeColor.FromRgb(22, 163, 74)), PdfAlign.Left, global::OfficeIMO.OfficeColor.FromRgb(22, 101, 52)))
                                    .PercentColumn(25, column => column.PanelParagraph(p => p.FontSize(global::OfficeIMO.OfficeLength.Points(22)).Bold("1.8 h").FontSize(global::OfficeIMO.OfficeLength.Points(9.5)).Text("\nMean first response"), MetricPanel(global::OfficeIMO.OfficeColor.FromRgb(239, 246, 255), global::OfficeIMO.OfficeColor.FromRgb(37, 99, 235)), PdfAlign.Left, global::OfficeIMO.OfficeColor.FromRgb(30, 64, 175)))
                                    .PercentColumn(25, column => column.PanelParagraph(p => p.FontSize(global::OfficeIMO.OfficeLength.Points(22)).Bold("37").FontSize(global::OfficeIMO.OfficeLength.Points(9.5)).Text("\nOpen escalations"), MetricPanel(global::OfficeIMO.OfficeColor.FromRgb(255, 251, 235), global::OfficeIMO.OfficeColor.FromRgb(217, 119, 6)), PdfAlign.Left, global::OfficeIMO.OfficeColor.FromRgb(146, 64, 14)))
                                    .PercentColumn(25, column => column.PanelParagraph(p => p.FontSize(global::OfficeIMO.OfficeLength.Points(22)).Bold("0").FontSize(global::OfficeIMO.OfficeLength.Points(9.5)).Text("\nCritical blockers"), MetricPanel(global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252), global::OfficeIMO.OfficeColor.FromRgb(100, 116, 139)), PdfAlign.Left, global::OfficeIMO.OfficeColor.FromRgb(51, 65, 85)));
                            });

                            content.Spacer(global::OfficeIMO.OfficeLength.Points(12));

                            // Chart and table on the left, narrative and decisions on the right.
                            content.Row(row => {
                                row.Gap(global::OfficeIMO.OfficeLength.Points(18))
                                    .PercentColumn(58, column => {
                                        column.Paragraph(p => p.Bold("Tickets resolved per month"), style: SectionLabel());
                                        column.Drawing(CreateTrendDrawing(), PdfAlign.Left, spacingBefore: global::OfficeIMO.OfficeLength.Points(2), spacingAfter: global::OfficeIMO.OfficeLength.Points(8));
                                        column.Table(CreateServiceRows(), style: ServiceTableStyle());
                                    })
                                    .PercentColumn(42, column => {
                                        column.Paragraph(p => p.Bold("Summary"), style: SectionLabel());
                                        column.PanelParagraph(
                                            p => p.Text("SLA attainment finished above the 92% target in all three months while resolved volume grew 15%. Integrations is the only service line below target, and its response time is the main driver."),
                                            new PdfPanelStyle {
                                                Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                                                BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225),
                                                BorderWidth = global::OfficeIMO.OfficeLength.Points(0.7),
                                                PaddingX = global::OfficeIMO.OfficeLength.Points(10),
                                                PaddingY = global::OfficeIMO.OfficeLength.Points(8)
                                            });
                                        column.Bullets(new[] {
                                            "Resolved tickets: 2,617, up from 2,274 last quarter.",
                                            "Security and Billing answer in about an hour.",
                                            "Integrations answers in 3.1 h against a 2 h goal.",
                                            "No critical blockers are open at quarter end."
                                        }, style: new PdfListStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(4), ItemSpacing = global::OfficeIMO.OfficeLength.Points(2) });
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

        private static PdfPanelStyle MetricPanel(global::OfficeIMO.OfficeColor background, global::OfficeIMO.OfficeColor border) {
            return new PdfPanelStyle {
                Background = background,
                BorderColor = border,
                BorderWidth = global::OfficeIMO.OfficeLength.Points(0.8),
                PaddingX = global::OfficeIMO.OfficeLength.Points(10),
                PaddingY = global::OfficeIMO.OfficeLength.Points(8)
            };
        }

        private static PdfParagraphStyle SectionLabel() {
            return new PdfParagraphStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(4), LineHeight = 1.1 };
        }

        // The chart fills the 58 % column so its right edge lines up with the table below:
        // (842 pt landscape A4 - 84 pt margins - 18 pt gap) x 58 %.
        private const double TrendChartWidth = (842D - 84D - 18D) * 0.58D;

        private static OfficeDrawing CreateTrendDrawing() {
            // The shared chart renderer gives round axis steps, readable labels and no decoration.
            var data = new OfficeChartData(
                new[] { "July", "August", "September" },
                new[] {
                    new OfficeSeries("Resolved", new[] { 812D, 871D, 934D }, null, OfficeColor.FromRgb(37, 99, 235)),
                    new OfficeSeries("Target", new[] { 850D, 850D, 850D }, null, OfficeColor.FromRgb(15, 23, 42), null,
                        showMarkers: false, strokeWidth: OfficeLength.Points(1.2D), strokeDashStyle: OfficeStrokeDashStyle.Dash,
                        renderKind: OfficeChartType.Line)
                });
            var style = new OfficeChartStyle(
                fontFamily: "Helvetica",
                backgroundColor: OfficeColor.White,
                showBorder: false,
                gridLineColor: OfficeColor.FromRgb(226, 232, 240),
                gridLineWidth: 0.5D,
                textColor: OfficeColor.FromRgb(30, 41, 59),
                mutedTextColor: OfficeColor.FromRgb(100, 116, 139),
                valueAxisColor: OfficeColor.Transparent,
                categoryAxisColor: OfficeColor.FromRgb(203, 213, 225))
                // Values sit in white at the foot of each bar, clear of the dashed target line across the bar tops.
                .WithDataLabelTextColor(OfficeColor.White);
            var layout = new OfficeChartLayout(
                legendPosition: OfficeChartLegendPosition.Top,
                showDataLabels: true,
                showDataLabelValues: true,
                dataLabelSeriesIndexes: new[] { 0 },
                dataLabelPosition: OfficeChartDataLabelPosition.InsideBase,
                verticalAxisNumberFormat: "#,##0");
            var chart = new OfficeChartSnapshot("Tickets resolved per month", null, OfficeChartType.ColumnClustered, data,
                width: OfficeLength.Points(TrendChartWidth), height: OfficeLength.Points(180D), style: style, layout: layout);
            return OfficeChartDrawingRenderer.Render(chart, useMinimumCanvas: false);
        }

        private static PdfTableCell[][] CreateServiceRows() {
            OfficeThemePalette palette = OfficeTheme.Default.Palette;
            PdfTableCell Status(string status) => PdfTableCell.RichTextCell(new[] {
                PdfTextRun.Bolded(status, status == "Watch" ? palette.Warning : palette.Success)
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
            var style = PdfTheme.FromOfficeTheme(OfficeTheme.Default).TableStyle!;
            style.CellPaddingX = OfficeLength.Points(7);
            style.CellPaddingY = OfficeLength.Points(5);
            style.CellPaddingTop = OfficeLength.Points(5);
            style.CellPaddingBottom = OfficeLength.Points(5);
            style.FontSize = OfficeLength.Points(9);
            style.RightAlignNumeric = true;
            style.ColumnWidthWeights = new List<double> { 1.6, 0.8, 1.0, 1.0, 1.1 };
            style.Alignments = new List<PdfColumnAlign> {
                    PdfColumnAlign.Left,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Left
                };
            return style;
        }

        private static PdfTableStyle DecisionTableStyle() {
            var style = PdfTheme.FromOfficeTheme(OfficeTheme.Default).TableStyle!;
            style.CellPaddingX = OfficeLength.Points(7);
            style.CellPaddingY = OfficeLength.Points(5);
            style.CellPaddingTop = OfficeLength.Points(5);
            style.CellPaddingBottom = OfficeLength.Points(5);
            style.HeaderRowCount = 1;
            style.FontSize = OfficeLength.Points(8.7);
            style.LineHeight = 1.15;
            style.ColumnWidthWeights = new List<double> { 1.0, 2.1 };
            return style;
        }
    }
}
