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
                        page.Size(PdfPageSizes.A4).Landscape().Margin(42, 38, 42, 42);
                        page.DefaultParagraphStyle(new PdfParagraphStyle { LineHeight = 1.18, SpacingAfter = 5 });
                        page.Content(content => {
                            content.Item(item => item.H1("Quarterly operations dashboard", PdfAlign.Left));
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

        // The chart fills the 58 % column so its right edge lines up with the table below:
        // (842 pt landscape A4 - 84 pt margins - 18 pt gap) x 58 %.
        private const double TrendChartWidth = (842D - 84D - 18D) * 0.58D;

        private static OfficeDrawing CreateTrendDrawing() {
            // The shared chart renderer gives round axis steps, readable labels and no decoration.
            var data = new OfficeChartData(
                new[] { "July", "August", "September" },
                new[] {
                    new OfficeChartSeries("Resolved", new[] { 812D, 871D, 934D }, null, OfficeColor.FromRgb(37, 99, 235)),
                    new OfficeChartSeries("Target", new[] { 850D, 850D, 850D }, null, OfficeColor.FromRgb(15, 23, 42), null,
                        showMarkers: false, strokeWidth: 1.2D, strokeDashStyle: OfficeStrokeDashStyle.Dash,
                        renderKind: OfficeChartKind.Line)
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
                categoryAxisColor: OfficeColor.FromRgb(203, 213, 225));
            var layout = new OfficeChartLayout(
                legendPosition: OfficeChartLegendPosition.Top,
                showDataLabels: true,
                showDataLabelValues: true,
                dataLabelSeriesIndexes: new[] { 0 },
                dataLabelPosition: OfficeChartDataLabelPosition.OutsideEnd,
                verticalAxisNumberFormat: "#,##0");
            var chart = new OfficeChartSnapshot("Tickets resolved per month", null, OfficeChartKind.ColumnClustered, data,
                widthPoints: TrendChartWidth, heightPoints: 180D, style: style, layout: layout);
            return OfficeChartDrawingRenderer.Render(chart, useMinimumCanvas: false);
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
            var style = PdfTheme.FromOfficeTheme(OfficeTheme.Default).TableStyle!;
            style.CellPaddingX = 7;
            style.CellPaddingY = 5;
            style.CellPaddingTop = 5;
            style.CellPaddingBottom = 5;
            style.FontSize = 9;
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
            style.CellPaddingX = 7;
            style.CellPaddingY = 5;
            style.CellPaddingTop = 5;
            style.CellPaddingBottom = 5;
            style.HeaderRowCount = 1;
            style.FontSize = 8.7;
            style.LineHeight = 1.15;
            style.ColumnWidthWeights = new List<double> { 1.0, 2.1 };
            return style;
        }
    }
}
