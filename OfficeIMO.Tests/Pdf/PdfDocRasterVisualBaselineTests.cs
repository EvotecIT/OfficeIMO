using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocRasterVisualBaselineTests {
    [Fact]
    public void ProfessionalReport_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("professional-report", CreateProfessionalReport);
    }

    [Fact]
    public void LineItemsTwoPage_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("line-items-two-page", CreateLineItemsTwoPage, pageCount: 2);
    }

    [Fact]
    public void HeadersFooters_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("headers-footers", CreateHeadersFooters, pageCount: 2);
    }

    [Fact]
    public void FlowDsl_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("flow-dsl", CreateFlowDsl, pageCount: 3);
    }

    [Theory]
    [InlineData("hello-world")]
    [InlineData("core-layout")]
    [InlineData("style-cheatsheet")]
    [InlineData("links-rules")]
    [InlineData("lists-tables")]
    [InlineData("table-style-gallery")]
    [InlineData("default-styles")]
    [InlineData("styled-runs")]
    [InlineData("drawing-gallery")]
    [InlineData("row-columns")]
    [InlineData("showcase-dashboard")]
    public void CorePdfScenarios_MatchPopplerRasterBaseline(string scenarioName) {
        AssertScenarioRasterBaseline(scenarioName, () => CreateCoreScenario(scenarioName));
    }

    private static void AssertScenarioRasterBaseline(string scenarioName, Func<byte[]> createPdf, int pageCount = 1) {
        if (pageCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageCount), pageCount, "Raster baseline page count must be positive.");
        }

        byte[] pdfBytes = createPdf();
        int actualPageCount = PdfInspector.Inspect(pdfBytes).PageCount;
        if (actualPageCount != pageCount) {
            throw new Xunit.Sdk.XunitException(
                "PDF raster baseline scenario '" + scenarioName + "' produced " +
                actualPageCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " page(s), but the approved baseline expects " +
                pageCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " page(s). Update the expected page count and baselines deliberately if this page-flow change is intended.");
        }

        if (!TryFindPdftoppm(out string rasterizerPath)) {
            if (IsRequired()) {
                throw new InvalidOperationException("PDF raster baseline tests require Poppler pdftoppm. Install Poppler or set OFFICEIMO_PDF_RASTERIZER to pdftoppm.exe.");
            }

            return;
        }

        string workDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.PdfRaster", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);
        string pdfPath = Path.Combine(workDir, scenarioName + ".pdf");

        try {
            File.WriteAllBytes(pdfPath, pdfBytes);
            for (int pageNumber = 1; pageNumber <= pageCount; pageNumber++) {
                string pageText = pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
                string outputPrefix = Path.Combine(workDir, scenarioName + "-page" + pageText);
                string actualPng = outputPrefix + ".png";

                RunPdftoppm(rasterizerPath, pdfPath, outputPrefix, workDir, pageNumber);

                if (!File.Exists(actualPng)) {
                    throw new FileNotFoundException("Poppler did not produce the expected PNG page snapshot.", actualPng);
                }

                AssertRasterBaseline("officeimo-pdf-" + scenarioName + ".page" + pageText + ".poppler.png", actualPng);
            }
        } finally {
            TryDeleteDirectory(workDir);
        }
    }

    [Fact]
    public void RasterBaseline_RejectsUnexpectedGeneratedPageCount() {
        var exception = Assert.Throws<Xunit.Sdk.XunitException>(() =>
            AssertScenarioRasterBaseline("page-count-mismatch", () =>
                PdfDoc.Create()
                    .Paragraph(p => p.Text("Page one"))
                    .PageBreak()
                    .Paragraph(p => p.Text("Page two"))
                    .ToBytes()));

        Assert.Contains("produced 2 page(s), but the approved baseline expects 1 page(s)", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RasterComparison_ReportsPixelDiffAndProducesDiffPng() {
        byte[] expected = PngRaster.EncodeRgb(2, 1, new byte[] {
            255, 255, 255,
            0, 0, 0
        });
        byte[] actual = PngRaster.EncodeRgb(2, 1, new byte[] {
            255, 255, 255,
            255, 0, 0
        });

        RasterComparison comparison = CompareRasterImages(expected, actual, channelTolerance: 0, allowedDifferentPixels: 0);

        Assert.False(comparison.Passed);
        Assert.Equal(1, comparison.DifferentPixels);
        Assert.Equal(2, comparison.TotalPixels);
        Assert.Equal(255, comparison.MaxChannelDelta);
        Assert.True(comparison.DiffPng.Length > 0);
        Assert.Equal(2, PngRaster.Decode(comparison.DiffPng).Width);
    }

    private static byte[] CreateCoreScenario(string scenarioName) {
        switch (scenarioName) {
            case "hello-world":
                return CreateHelloWorld();
            case "core-layout":
                return CreateCoreLayout();
            case "style-cheatsheet":
                return CreateStyleCheatsheet();
            case "links-rules":
                return CreateLinksAndRules();
            case "lists-tables":
                return CreateListsTables();
            case "table-style-gallery":
                return CreateTableStyleGallery();
            case "default-styles":
                return CreateDefaultStyles();
            case "styled-runs":
                return CreateStyledRuns();
            case "drawing-gallery":
                return CreateDrawingGallery();
            case "row-columns":
                return CreateRowColumns();
            case "showcase-dashboard":
                return CreateShowcaseDashboard();
            default:
                throw new ArgumentOutOfRangeException(nameof(scenarioName), scenarioName, "Unknown PDF raster scenario.");
        }
    }

    private static byte[] CreateHelloWorld() {
        return PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 11,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf hello-world - page {page}/{pages}",
                FooterAlign = PdfAlign.Center,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Hello World", author: "OfficeIMO")
            .H1("OfficeIMO.Pdf Hello World", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("This deterministic smoke page protects basic headings, paragraphs, metadata, and footer rendering."))
            .Paragraph(p => p.Text("It is intentionally small so visual regressions in the simplest document path are easy to spot."))
            .ToBytes();
    }

    private static byte[] CreateCoreLayout() {
        return PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf core layout - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Core Layout", author: "OfficeIMO")
            .H1("Core Layout Baseline", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p
                .Text("This page mixes ")
                .Bold("rich text")
                .Text(", colors, list flow, a panel, a horizontal rule, and a wrapped table."),
                style: new PdfParagraphStyle { LineHeight = 1.5, LeftIndent = 6, RightIndent = 18, SpacingAfter = 7 })
            .Bullets(new[] {
                "Natural proportional text spacing",
                "Stable vertical rhythm",
                "Readable table cells"
            }, color: PdfColor.FromRgb(55, 65, 81))
            .Numbered(new[] {
                "Measure the generated document.",
                "Compare the approved raster image.",
                "Improve the engine instead of hiding rough output."
            }, color: PdfColor.FromRgb(55, 65, 81))
            .PanelParagraph(
                p => p.Bold("Panel check").LineBreak().Text("Text should sit comfortably inside its border and retain readable padding after rasterization."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(248, 250, 252),
                    BorderColor = PdfColor.FromRgb(183, 194, 207),
                    PaddingX = 9,
                    PaddingY = 7
                },
                defaultColor: PdfColor.FromRgb(35, 88, 65))
            .HR(thickness: 1, color: PdfColor.FromRgb(32, 76, 120), spacingBefore: 8, spacingAfter: 8)
            .Table(new[] {
                new[] { "Area", "Expectation", "State" },
                new[] { "Paragraphs", "Rich text wraps and keeps comfortable line spacing.", "Guarded" },
                new[] { "Lists", "Bullets and numbered steps align with wrapped text.", "Guarded" },
                new[] { "Tables", "Long text stays inside cells and body text remains readable after colored headers.", "Guarded" }
            }, style: CreateCoreTableStyle())
            .Paragraph(p => p.Text("End of core layout baseline."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80))
            .ToBytes();
    }

    private static byte[] CreateListsTables() {
        var rows = new[] {
            new[] { "Item", "Qty", "Unit", "Total", "Notes" },
            new[] { "Monitoring seats", "25", "$4.50", "$112.50", "Monthly report-ready subscription line." },
            new[] { "Security review", "1", "$250.00", "$250.00", "Includes a short executive summary and remediation list." },
            new[] { "Documentation pack", "3", "$35.00", "$105.00", "Generated attachments for PSWriteOffice hand-off." },
            new[] { "Total", "", "", "$467.50", "Ready for approval." }
        };

        return PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf lists and tables",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Lists and Tables", author: "OfficeIMO")
            .H1("Lists and Tables", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("A compact report sample for list rhythm, numeric alignment, footer rows, and wrapped table notes."))
            .Bullets(new[] {
                "Report-friendly list spacing",
                "Right-aligned quantities and amounts",
                "Footer row that remains visually distinct"
            }, PdfAlign.Left, PdfColor.FromRgb(55, 65, 81))
            .Numbered(new[] {
                "Collect line items.",
                "Render the table with stable column widths.",
                "Review the generated PDF through the raster gate."
            }, PdfAlign.Left, PdfColor.FromRgb(55, 65, 81))
            .Table(rows, PdfAlign.Left, CreateLineItemTableStyle())
            .Paragraph(p => p.Text("End of lists and tables sample."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80))
            .ToBytes();
    }

    private static byte[] CreateDefaultStyles() {
        var rows = new[] {
            new[] { "Metric", "Current", "Target" },
            new[] { "Runtime dependencies", "0", "0" },
            new[] { "Visual gates", "Growing", "Required for public claims" },
            new[] { "PowerShell wrapper", "PSWriteOffice", "Expose safe PDF operations" }
        };

        return PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf default styles",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true,
                DefaultTableStyle = TableStyles.Light()
            })
            .Meta(title: "OfficeIMO.Pdf Default Styles", author: "OfficeIMO")
            .H1("Default Styles", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("This sample uses document-level defaults for text color, headers, footers, and the light table preset."))
            .PanelParagraph(
                p => p.Text("The default table style should be good enough for a simple business report without every caller hand-tuning colors and padding."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(248, 250, 252),
                    BorderColor = PdfColor.FromRgb(183, 194, 207),
                    PaddingX = 9,
                    PaddingY = 7
                })
            .Table(rows)
            .ToBytes();
    }

    private static byte[] CreateStyledRuns() {
        return PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf styled runs",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Styled Runs", author: "OfficeIMO")
            .H1("Styled Runs", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("A compact visual sample for inline font style, color, underline, strike-through, and stateful run toggles."))
            .PanelParagraph(
                p => p
                    .Bold("Inline styles")
                    .Text(" should remain readable in normal business-report text, not only in synthetic text extraction checks."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(248, 250, 252),
                    BorderColor = PdfColor.FromRgb(183, 194, 207),
                    PaddingX = 9,
                    PaddingY = 7
                })
            .Paragraph(p => p
                .Text("You can mix ")
                .Bold("bold ")
                .Italic("italic ")
                .Bold(true).Italic(true).Text("bold italic ").Italic(false).Bold(false)
                .Underlined("underlined ")
                .Strikethrough("obsolete ")
                .Color(PdfColor.FromRgb(80, 80, 80)).Text("and ")
                .Color(PdfColor.FromRgb(8, 28, 120)).Text("colors."),
                style: new PdfParagraphStyle { SpacingBefore = 14 })
            .Paragraph(p => p
                .Text("Underline respects color: ")
                .Underlined("red", PdfColor.FromRgb(200, 0, 0))
                .Text(", ")
                .Underlined("blue", PdfColor.FromRgb(20, 90, 180))
                .Text(", and ")
                .Underlined("green", PdfColor.FromRgb(0, 128, 0))
                .Text("."))
            .Paragraph(p => p
                .Text("Stateful color toggles: ")
                .Color(PdfColor.FromRgb(185, 28, 28)).Text("critical ")
                .Color(PdfColor.FromRgb(20, 90, 180)).Text("informational ")
                .Color(PdfColor.FromRgb(22, 101, 52)).Text("healthy ")
                .Color(PdfColor.FromRgb(31, 41, 55)).Text("normal."))
            .Paragraph(p => p.Text("End of styled runs sample."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80))
            .ToBytes();
    }

    private static byte[] CreateFlowDsl() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10,
            DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
            HeaderFont = PdfStandardFont.Helvetica,
            HeaderFontSize = 8,
            FooterFont = PdfStandardFont.Helvetica,
            FooterFontSize = 8
        };

        return PdfDoc.Create(options)
            .Compose(document => {
                document.Page(page => {
                    page.Size(PageSizes.A5);
                    page.Margin(36, 42, 36, 42);
                    page.DefaultTextStyle(x => x
                        .Font(PdfStandardFont.Helvetica)
                        .FontSize(10)
                        .Color(PdfColor.FromRgb(31, 41, 55)));
                    page.Header(h => h.AlignLeft().Text("OfficeIMO.Pdf compose DSL gate"));

                    page.Content(c => c
                        .Column(column => {
                            column.Item().H1("Compose DSL");
                            column.Item().Paragraph(p => p.Text("A compact multi-page visual baseline for the OfficeIMO.Pdf composition API."));
                            column.Item().PanelParagraph(
                                p => p
                                    .Bold("What this protects")
                                    .LineBreak()
                                    .Text("Page settings, composed content, explicit page breaks, header/footer tokens, and rich text inside composed items."),
                                new PanelStyle {
                                    Background = PdfColor.FromRgb(248, 250, 252),
                                    BorderColor = PdfColor.FromRgb(183, 194, 207),
                                    PaddingX = 9,
                                    PaddingY = 7
                                });
                            column.Item().PageBreak();

                            column.Item().H2("Operational Notes");
                            column.Item().Bullets(new[] {
                                "Compose uses the same document engine as fluent blocks.",
                                "Visual gates should cover both public authoring styles.",
                                "Footer page totals must remain correct after explicit breaks."
                            }, color: PdfColor.FromRgb(55, 65, 81));
                            column.Item().HR(0.8, PdfColor.FromRgb(183, 194, 207), 8, 8);
                            column.Item().Paragraph(p => p
                                .Text("Status: ")
                                .Bold("ready for wrapper experiments ")
                                .Color(PdfColor.FromRgb(20, 90, 180)).Text("once visual quality keeps improving."));
                            column.Item().PageBreak();

                            column.Item().H2("Color Sections");
                            column.Item().Paragraph(p => p.Color(PdfColor.FromRgb(185, 28, 28)).Text("Critical items should remain readable when colored inline."));
                            column.Item().Paragraph(p => p.Color(PdfColor.FromRgb(22, 101, 52)).Text("Healthy items should use calm green without overpowering the page."));
                            column.Item().Paragraph(p => p.Color(PdfColor.FromRgb(20, 90, 180)).Text("Informational items should keep good contrast on white backgrounds."));
                            column.Item().Paragraph(p => p.Text("End of compose DSL sample."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80));
                        }));

                    page.Footer(f => f.AlignCenter().Text(t => t
                        .Text("OfficeIMO.Pdf compose - page ")
                        .CurrentPage()
                        .Text("/")
                        .TotalPages()));
                });
            })
            .ToBytes();
    }

    internal static byte[] CreateShowcaseDashboard() {
        return PdfDoc.Create(new PdfOptions {
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
                                .Column(25, column => column.PanelParagraph(p => p.Bold("92%").Text("\nSLA attainment"), CreateDashboardMetricPanel(PdfColor.FromRgb(236, 253, 245), PdfColor.FromRgb(22, 163, 74)), PdfAlign.Left, PdfColor.FromRgb(22, 101, 52)))
                                .Column(25, column => column.PanelParagraph(p => p.Bold("1.8h").Text("\nMean response"), CreateDashboardMetricPanel(PdfColor.FromRgb(239, 246, 255), PdfColor.FromRgb(37, 99, 235)), PdfAlign.Left, PdfColor.FromRgb(30, 64, 175)))
                                .Column(25, column => column.PanelParagraph(p => p.Bold("34").Text("\nOpen actions"), CreateDashboardMetricPanel(PdfColor.FromRgb(255, 251, 235), PdfColor.FromRgb(217, 119, 6)), PdfAlign.Left, PdfColor.FromRgb(146, 64, 14)))
                                .Column(25, column => column.PanelParagraph(p => p.Bold("0").Text("\nCritical blockers"), CreateDashboardMetricPanel(PdfColor.FromRgb(248, 250, 252), PdfColor.FromRgb(100, 116, 139)), PdfAlign.Left, PdfColor.FromRgb(51, 65, 85)));
                        });

                        content.Spacer(12);
                        content.Row(row => {
                            row.Gap(18)
                                .Column(58, column => {
                                    column.Paragraph(p => p.Bold("Delivery trend"), style: CreateDashboardSectionLabelStyle());
                                    column.Drawing(CreateDashboardTrendDrawing(), PdfAlign.Left, spacingBefore: 2, spacingAfter: 8);
                                    column.Table(CreateDashboardRiskRows(), style: CreateDashboardRiskTableStyle());
                                })
                                .Column(42, column => {
                                    column.Paragraph(p => p.Bold("Narrative"), style: CreateDashboardSectionLabelStyle());
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
                                    column.Table(CreateDashboardDecisionRows(), style: CreateDashboardDecisionTableStyle());
                                });
                        });
                    });
                });
            })
            .ToBytes();
    }

    private static PanelStyle CreateDashboardMetricPanel(PdfColor background, PdfColor border) {
        return new PanelStyle {
            Background = background,
            BorderColor = border,
            BorderWidth = 0.8,
            PaddingX = 10,
            PaddingY = 8
        };
    }

    private static PdfParagraphStyle CreateDashboardSectionLabelStyle() {
        return new PdfParagraphStyle { SpacingBefore = 0, SpacingAfter = 4, LineHeight = 1.1 };
    }

    private static OfficeDrawing CreateDashboardTrendDrawing() {
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

    private static string[][] CreateDashboardRiskRows() {
        return new[] {
            new[] { "Area", "State", "Trend", "Owner" },
            new[] { "PDF layout rhythm", "Good", "+12%", "OfficeIMO.Pdf" },
            new[] { "Table wrapping", "Watch", "-3%", "Renderer" },
            new[] { "Read/manipulation", "Growing", "+31%", "Core" },
            new[] { "Word/Excel export path", "Planned", "+0%", "Roadmap" }
        };
    }

    private static PdfTableCell[][] CreateDashboardDecisionRows() {
        return new[] {
            new[] { PdfTableCell.Span("Next decisions", 2) },
            new[] { PdfTableCell.TextCell("Visual fixtures"), PdfTableCell.TextCell("Keep generic, use documents as gates") },
            new[] { PdfTableCell.TextCell("AST model"), PdfTableCell.TextCell("Promote page/content tree over helper-only APIs") },
            new[] { PdfTableCell.TextCell("Conversion"), PdfTableCell.TextCell("Add Word/Excel/PPT exporters in slices") }
        };
    }

    private static byte[] CreateDrawingGallery() {
        return PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Drawing shared vector gate",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Drawing Gallery", author: "OfficeIMO")
            .H1("Drawing Gallery", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("A visual baseline for shared OfficeIMO.Drawing vector descriptors rendered by the dependency-free PDF engine."))
            .Drawing(CreateDrawingScene(), PdfAlign.Center, spacingBefore: 8, spacingAfter: 10)
            .Paragraph(p => p
                .Text("Covered: ")
                .Bold("gradients")
                .Text(", shadows, dashed strokes, line caps and joins, clipping paths, transforms, grouped scenes, and freeform paths."))
            .ToBytes();
    }

    private static OfficeDrawing CreateDrawingScene() {
        var drawing = new OfficeDrawing(420, 170);

        var background = OfficeShape.RoundedRectangle(420, 170, 12);
        background.FillColor = OfficeColor.FromRgb(248, 250, 252);
        background.StrokeColor = OfficeColor.FromRgb(183, 194, 207);
        background.StrokeWidth = 0.8;
        drawing.AddShape(background, 0, 0);

        var ribbon = OfficeShape.RoundedRectangle(132, 42, 10);
        ribbon.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.FromRgb(25, 55, 85), OfficeColor.FromRgb(20, 90, 180));
        ribbon.StrokeColor = OfficeColor.FromRgb(25, 55, 85);
        ribbon.StrokeWidth = 0.8;
        ribbon.Shadow = new OfficeShadow(OfficeColor.FromRgb(15, 23, 42), 0.18, 3, 3);
        drawing.AddShape(ribbon, 18, 18);

        var ellipse = OfficeShape.Ellipse(64, 42);
        ellipse.FillColor = OfficeColor.FromRgb(220, 252, 231);
        ellipse.StrokeColor = OfficeColor.FromRgb(22, 101, 52);
        ellipse.StrokeWidth = 1.2;
        drawing.AddShape(ellipse, 176, 18);

        var triangle = OfficeShape.Polygon(new OfficePoint(0, 38), new OfficePoint(36, 0), new OfficePoint(72, 38));
        triangle.FillColor = OfficeColor.FromRgb(254, 243, 199);
        triangle.StrokeColor = OfficeColor.FromRgb(180, 83, 9);
        triangle.StrokeWidth = 1.2;
        triangle.StrokeLineJoin = OfficeStrokeLineJoin.Round;
        drawing.AddShape(triangle, 276, 20);

        var rule = OfficeShape.Line(0, 0, 380, 0);
        rule.StrokeColor = OfficeColor.FromRgb(80, 80, 80);
        rule.StrokeWidth = 1.4;
        rule.StrokeDashStyle = OfficeStrokeDashStyle.DashDot;
        rule.StrokeLineCap = OfficeStrokeLineCap.Round;
        drawing.AddShape(rule, 20, 84);

        var clipped = OfficeShape.Rectangle(76, 44);
        clipped.FillColor = OfficeColor.FromRgb(219, 234, 254);
        clipped.StrokeColor = OfficeColor.FromRgb(20, 90, 180);
        clipped.StrokeWidth = 1;
        clipped.ClipPath = OfficeClipPath.RoundedRectangle(76, 44, 12);
        drawing.AddShape(clipped, 26, 110);

        var transformed = OfficeShape.Rectangle(74, 34);
        transformed.FillColor = OfficeColor.FromRgb(237, 233, 254);
        transformed.StrokeColor = OfficeColor.FromRgb(91, 33, 182);
        transformed.StrokeWidth = 1;
        transformed.Transform = OfficeTransform.RotateDegrees(8, 37, 17);
        drawing.AddShape(transformed, 132, 114);

        var path = OfficeShape.Path(
            OfficePathCommand.MoveTo(0, 34),
            OfficePathCommand.CubicBezierTo(22, -8, 60, -8, 82, 34),
            OfficePathCommand.LineTo(82, 44),
            OfficePathCommand.LineTo(0, 44),
            OfficePathCommand.Close());
        path.FillColor = OfficeColor.FromRgb(252, 231, 243);
        path.StrokeColor = OfficeColor.FromRgb(157, 23, 77);
        path.StrokeWidth = 1;
        drawing.AddShape(path, 238, 108);

        return drawing;
    }

    private static byte[] CreateRowColumns() {
        return PdfDoc.Create(new PdfOptions {
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
            .ToBytes();
    }

    private static byte[] CreateHeadersFooters() {
        return PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf header/footer gate",
                HeaderAlign = PdfAlign.Left,
                HeaderOffsetY = 18,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Center,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Headers and Footers", author: "OfficeIMO")
            .H1("Header and Footer Baseline", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("Page one protects header placement, footer placement, and page number rendering on the first page."))
            .PanelParagraph(
                p => p.Text("The same options should continue to render consistently after an explicit page break."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(248, 250, 252),
                    BorderColor = PdfColor.FromRgb(183, 194, 207),
                    PaddingX = 9,
                    PaddingY = 7
                })
            .PageBreak()
            .H2("Continuation Page", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("Page two proves the raster harness compares more than the first page and guards the {page}/{pages} footer tokens."))
            .Paragraph(p => p.Text("Right-aligned continuation note."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80))
            .ToBytes();
    }

    private static byte[] CreateStyleCheatsheet() {
        return PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf style cheatsheet",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Style Cheatsheet", author: "OfficeIMO")
            .H1("Style Cheatsheet", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .PanelParagraph(
                p => p.Text("A compact visual sample for rich text, color, underline, and alignment behavior."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(248, 250, 252),
                    BorderColor = PdfColor.FromRgb(183, 194, 207),
                    PaddingX = 9,
                    PaddingY = 7
                })
            .Paragraph(p => p
                .Text("Normal ")
                .Bold("Bold ")
                .Italic("Italic ")
                .Bold("Bold").Italic(" Italic ")
                .Underlined("Underline "))
            .Paragraph(p => p
                .Text("Colors: ")
                .Color(PdfColor.FromRgb(200, 0, 0)).Text("Red ")
                .Color(PdfColor.FromRgb(20, 90, 180)).Text("Blue ")
                .Color(PdfColor.FromRgb(0, 128, 0)).Text("Green"))
            .Paragraph(p => p
                .Text("Combinations: ")
                .Bold("Bold ")
                .Italic("Italic ")
                .Underlined("Underlined ")
                .Bold("Bold ").Italic("Italic ").Underlined("Underlined "))
            .Paragraph(p => p
                .Text("Stateful toggles: ")
                .Bold(true).Text("bold on ")
                .Bold(false).Text("bold off ")
                .Italic(true).Text("italic on ")
                .Italic(false).Text("italic off ")
                .Underline(true).Text("ul on ")
                .Underline(false).Text("ul off"))
            .HR(0.8, PdfColor.FromRgb(183, 194, 207), 8, 8)
            .Paragraph(p => p.Text("Center aligned line"), PdfAlign.Center)
            .Paragraph(p => p.Text("Right aligned line"), PdfAlign.Right)
            .ToBytes();
    }

    private static byte[] CreateTableStyleGallery() {
        var rows = new[] {
            new[] { "Signal", "State", "Notes" },
            new[] { "Header", "Repeated", "The first row should stay readable without relying on a domain-specific preset." },
            new[] { "Flow", "Generic", "Borders, row separators, and spacing should reveal the preset shape at raster level." }
        };

        PdfDoc doc = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 9.5,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf Word-like table styles",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Table Style Gallery", author: "OfficeIMO")
            .H1("Table Style Gallery", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("Generic Word-style table names rendered by OfficeIMO.Pdf without invoice or report-specific behavior."));

        foreach (string styleName in TableStyles.SupportedWordStyleNames) {
            PdfTableStyle style = TableStyles.FromWordTableStyle(styleName);
            style.Caption = styleName;
            style.CaptionColor = PdfColor.FromRgb(80, 90, 100);
            style.CaptionFontSize = 8.5;
            style.CaptionSpacingAfter = 4;
            style.SpacingBefore = 6;
            style.SpacingAfter = 6;
            style.ColumnWidthPoints = new List<double?> { 76, 70, 300 };
            style.AutoFitColumns = false;
            style.Alignments = new List<PdfColumnAlign> {
                PdfColumnAlign.Left,
                PdfColumnAlign.Center,
                PdfColumnAlign.Left
            };

            doc.Table(rows, PdfAlign.Left, style);
        }

        return doc.ToBytes();
    }

    private static byte[] CreateLinksAndRules() {
        var rows = new[] {
            new[] { "Site", "Label", "Notes" },
            new[] { "OfficeIMO", "Homepage", "Docs" },
            new[] { "GitHub", "Repo", "Issues" }
        };

        var links = new Dictionary<(int Row, int Col), string> {
            [(1, 0)] = "https://officeimo.net/",
            [(1, 1)] = "https://officeimo.net/",
            [(1, 2)] = "https://officeimo.net/docs",
            [(2, 0)] = "https://github.com/EvotecIT/OfficeIMO",
            [(2, 1)] = "https://github.com/EvotecIT/OfficeIMO",
            [(2, 2)] = "https://github.com/EvotecIT/OfficeIMO/issues"
        };

        return PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf links and rules",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Links and Rules", author: "OfficeIMO")
            .H1("Links & Rules Demo", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85), linkUri: "https://github.com/EvotecIT/OfficeIMO")
            .Paragraph(p => p
                .Text("Visit ")
                .Link("OfficeIMO GitHub", "https://github.com/EvotecIT/OfficeIMO", PdfColor.FromRgb(20, 90, 180))
                .Text(" and the ")
                .Link("project website", "https://officeimo.net/", PdfColor.FromRgb(20, 90, 180))
                .Text(" for more details."))
            .HR(0.8, PdfColor.FromRgb(183, 194, 207), 8, 8)
            .TableWithLinks(rows, links, PdfAlign.Left, CreateLinksTableStyle())
            .PanelParagraph(
                p => p
                    .Text("You can also place links ")
                    .Link("inside panels", "https://officeimo.net/", PdfColor.FromRgb(20, 90, 180))
                    .Text("."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(248, 250, 252),
                    BorderColor = PdfColor.FromRgb(183, 194, 207),
                    BorderWidth = 0.5,
                    PaddingX = 9,
                    PaddingY = 7
                })
            .ToBytes();
    }

    private static byte[] CreateProfessionalReport() {
        string logoPath = Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png");
        byte[] logo = File.Exists(logoPath) ? File.ReadAllBytes(logoPath) : CreateFallbackLogo();
        byte[] alphaBadge = CreateTransparentBadgePng();

        return PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf professional report",
                HeaderAlign = PdfAlign.Left,
                HeaderOffsetY = 18,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf professional report - page {page}/{pages}",
                FooterAlign = PdfAlign.Center,
                ShowPageNumbers = true,
                CreateOutlineFromHeadings = true
            })
            .Meta(title: "OfficeIMO.Pdf Professional Report", author: "OfficeIMO")
            .H1("Executive Security Summary", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p
                .Text("A polished first-party PDF report built with ")
                .Bold("OfficeIMO.Pdf")
                .Text(", shared drawing descriptors, wrapped tables, images, and generated bookmarks."))
            .Image(logo, 120, 48, PdfAlign.Right, fit: OfficeImageFit.Contain)
            .PanelParagraph(
                p => p.Bold("Healthy").Text(" posture with no critical findings."),
                CreateStatusPanelStyle(),
                defaultColor: PdfColor.FromRgb(42, 132, 82))
            .Image(alphaBadge, 18, 18, PdfAlign.Left)
            .Shape(CreateAccentRibbon(), spacingBefore: 12, spacingAfter: 10)
            .PanelParagraph(
                p => p.Text("Operator note: long values should wrap cleanly, tables should stay inside the page, and reusable drawing primitives should remain available to Word, Excel, PowerPoint, and PDF exporters."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(248, 250, 252),
                    BorderColor = PdfColor.FromRgb(183, 194, 207),
                    PaddingX = 9,
                    PaddingY = 7
                })
            .Table(new[] {
                new[] { "Signal", "Evidence", "Action" },
                new[] { "DMARC", "Policy is enforced for the primary domain and aligned subdomains.", "Monitor" },
                new[] { "TLS", "Certificate chain and protocol posture are ready for automated PSWriteOffice reports.", "Keep" },
                new[] { "DNS", "Delegation and stale-record checks are summarized without overflowing table cells.", "Review" },
                new[] { "PDF", "The report is generated by the MIT licensed dependency-free OfficeIMO.Pdf engine.", "Expand" }
            }, style: CreateReportTableStyle())
            .Image(logo, 112, 36, PdfAlign.Center, fit: OfficeImageFit.Contain)
            .Paragraph(p => p.Text("Generated by OfficeIMO.Pdf."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80))
            .ToBytes();
    }

    private static byte[] CreateLineItemsTwoPage() {
        string logoPath = Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png");
        byte[] logo = File.Exists(logoPath) ? File.ReadAllBytes(logoPath) : CreateFallbackLogo();
        var lineItemRows = CreateLineItemRows();
        var lineItemStyle = CreateLineItemGateTableStyle();
        var totalsRows = new[] {
            new[] { "Subtotal", "5 201,32 PLN" },
            new[] { "VAT 23%", "1 196,30 PLN" },
            new[] { "Total", "6 397,62 PLN" }
        };

        return PdfDoc.Create(new PdfOptions {
                PageWidth = 595,
                PageHeight = 842,
                MarginLeft = 50,
                MarginRight = 50,
                MarginTop = 54,
                MarginBottom = 58,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 11,
                DefaultTextColor = PdfColor.FromRgb(25, 25, 25),
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 11,
                FooterFormat = "{page} / {pages}",
                FooterAlign = PdfAlign.Center,
                FooterOffsetY = 28,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Line-Item Visual Gate", author: "OfficeIMO")
            .Compose(document => {
                document.Page(page => {
                    page.Content(content => {
                        content.Row(row => row
                            .Column(58, column => column
                                .H1("Statement #4048")
                                .Paragraph(p => p
                                    .Text("Prepared: 23/12/2025\n")
                                    .Text("Review by: 06/01/2026"),
                                    style: new PdfParagraphStyle { SpacingAfter = 2 }))
                            .Column(42, column => column
                                .Image(logo, 156, 54, PdfAlign.Right, fit: OfficeImageFit.Contain)));

                        content.Column(column => column
                            .Item()
                            .Spacer(41.4));

                        content.Row(row => row
                            .Gap(30)
                            .Column(47, column => column
                                .H3("Prepared by")
                                .HR(1.2, PdfColor.Black, spacingBefore: 2, spacingAfter: 10)
                                .Paragraph(p => p.Text("Syllabae Repraesentant\nOvum picem\nPrinceps avem distant, Linteum amicitia\nofficium21@aut statum.com\n881-306-3914"),
                                    style: new PdfParagraphStyle { LineHeight = 1.25, SpacingAfter = 0 }))
                            .Column(47, column => column
                                .H3("Recipient")
                                .HR(1.2, PdfColor.Black, spacingBefore: 2, spacingAfter: 10)
                                .Paragraph(p => p.Text("Ceciderit Original\nAurum currunt\nSolis multum platea, Cocus fuge fluvio\nsubsisto93@celeritate.com\n839-621-9110"),
                                    style: new PdfParagraphStyle { LineHeight = 1.25, SpacingAfter = 0 })));

                        content.Column(column => {
                            column.Item()
                                .Spacer(33.4)
                                .Table(lineItemRows, PdfAlign.Left, lineItemStyle)
                                .Table(totalsRows, PdfAlign.Right, CreateLineItemTotalsTableStyle())
                                .PanelParagraph(
                                    p => p.Bold("Document note: ").Text("Project details, approval notes, or wrapper-provided metadata can be placed here by PSWriteOffice."),
                                    new PanelStyle {
                                        Background = PdfColor.FromRgb(248, 250, 252),
                                        BorderColor = PdfColor.FromRgb(210, 218, 226),
                                        BorderWidth = 0.5,
                                        PaddingX = 8,
                                        PaddingY = 6,
                                        SpacingBefore = 10,
                                        SpacingAfter = 0
                                    });
                        });
                    });
                });
            })
            .ToBytes();
    }

    private static PdfTableStyle CreateCoreTableStyle() {
        return new PdfTableStyle {
            HeaderFill = PdfColor.FromRgb(32, 76, 120),
            HeaderTextColor = PdfColor.White,
            TextColor = PdfColor.FromRgb(31, 41, 55),
            RowStripeFill = PdfColor.FromRgb(248, 250, 252),
            BorderColor = PdfColor.FromRgb(210, 218, 226),
            BorderWidth = 0.5,
            CellPaddingX = 6,
            CellPaddingY = 5,
            Caption = "Table 1. Core layout checks",
            CaptionColor = PdfColor.FromRgb(80, 90, 100),
            CaptionFontSize = 8.5,
            CaptionSpacingAfter = 5,
            SpacingBefore = 6,
            SpacingAfter = 14,
            ColumnWidthPoints = new List<double?> { 82, 310, 76 },
            AutoFitColumns = false
        };
    }

    private static List<string[]> CreateLineItemRows() {
        var rows = new List<string[]> {
            new[] { "#", "Product", "Unit price", "Quantity", "Total" }
        };
        string[] products = {
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
            "Umero certus tantum",
            "Aptent saluto",
            "Nubes vigil pretium",
            "Clarus vectigal",
            "Integer cursus",
            "Lacus civitas",
            "Sodalitas pretium",
            "Vestis angularis",
            "Fractus lumen",
            "Nomen porttitor",
            "Finis officium"
        };
        decimal[] prices = { 31.80m, 62.57m, 42.50m, 22.75m, 85.56m, 40.58m, 37.72m, 88.21m, 59.83m, 21.85m, 6.41m, 93.57m, 77.27m, 9.68m, 23.94m, 79.05m, 11.65m, 81.72m, 18.44m, 55.10m, 72.35m, 14.25m, 44.70m, 38.12m, 66.90m, 29.95m, 53.42m, 47.18m };
        int[] quantities = { 2, 7, 5, 5, 2, 7, 7, 5, 1, 9, 8, 7, 1, 8, 6, 8, 7, 2, 6, 3, 4, 9, 2, 5, 3, 7, 4, 6 };

        for (int i = 0; i < products.Length; i++) {
            decimal total = prices[i] * quantities[i];
            rows.Add(new[] {
                (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture),
                products[i],
                FormatPln(prices[i]),
                quantities[i].ToString(System.Globalization.CultureInfo.InvariantCulture),
                FormatPln(total)
            });
        }

        return rows;
    }

    private static PdfTableStyle CreateLineItemGateTableStyle() {
        var style = TableStyles.ListTable1Light();
        style.SpacingBefore = 0;
        style.SpacingAfter = 14;
        style.Alignments = new List<PdfColumnAlign> {
            PdfColumnAlign.Right,
            PdfColumnAlign.Left,
            PdfColumnAlign.Right,
            PdfColumnAlign.Right,
            PdfColumnAlign.Right
        };
        style.ColumnWidthPoints = new List<double?> { 28, 235, 82, 55, 95 };
        style.AutoFitColumns = false;
        return style;
    }

    private static PdfTableStyle CreateLineItemTotalsTableStyle() {
        return new PdfTableStyle {
            HeaderFill = null,
            HeaderTextColor = PdfColor.Black,
            TextColor = PdfColor.FromRgb(25, 25, 25),
            FooterFill = null,
            FooterTextColor = PdfColor.Black,
            RowStripeFill = null,
            BorderColor = PdfColor.FromRgb(224, 224, 224),
            BorderWidth = 0.4,
            CellPaddingX = 6,
            CellPaddingY = 5,
            HeaderRowCount = 0,
            FooterRowCount = 1,
            SpacingBefore = 0,
            SpacingAfter = 10,
            Alignments = new List<PdfColumnAlign> {
                PdfColumnAlign.Right,
                PdfColumnAlign.Right
            },
            ColumnWidthPoints = new List<double?> { 105, 105 }
        };
    }

    private static string FormatPln(decimal value) =>
        value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture).Replace('.', ',') + " PLN";

    private static PdfTableStyle CreateLineItemTableStyle() {
        return new PdfTableStyle {
            HeaderFill = PdfColor.FromRgb(32, 76, 120),
            HeaderTextColor = PdfColor.White,
            TextColor = PdfColor.FromRgb(31, 41, 55),
            FooterFill = PdfColor.FromRgb(232, 241, 248),
            FooterTextColor = PdfColor.FromRgb(25, 55, 85),
            RowStripeFill = PdfColor.FromRgb(248, 250, 252),
            BorderColor = PdfColor.FromRgb(210, 218, 226),
            BorderWidth = 0.5,
            CellPaddingX = 6,
            CellPaddingY = 5,
            HeaderRowCount = 1,
            FooterRowCount = 1,
            Caption = "Table 1. Example line items",
            CaptionColor = PdfColor.FromRgb(80, 90, 100),
            CaptionFontSize = 8.5,
            CaptionSpacingAfter = 5,
            SpacingBefore = 6,
            SpacingAfter = 14,
            RightAlignNumeric = true,
            Alignments = new List<PdfColumnAlign> {
                PdfColumnAlign.Left,
                PdfColumnAlign.Right,
                PdfColumnAlign.Right,
                PdfColumnAlign.Right,
                PdfColumnAlign.Left
            },
            ColumnWidthPoints = new List<double?> { 110, 38, 58, 68, 170 },
            AutoFitColumns = false
        };
    }

    private static PdfTableStyle CreateDashboardRiskTableStyle() {
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

    private static PdfTableStyle CreateDashboardDecisionTableStyle() {
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

    private static PanelStyle CreateStatusPanelStyle() {
        return new PanelStyle {
            Background = PdfColor.FromRgb(230, 247, 238),
            BorderColor = PdfColor.FromRgb(42, 132, 82),
            BorderWidth = 1.2,
            PaddingX = 8,
            PaddingY = 5,
            MaxWidth = 245
        };
    }

    private static OfficeShape CreateAccentRibbon() {
        var accent = OfficeShape.RoundedRectangle(168, 8, 4);
        accent.FillGradient = OfficeLinearGradient.Horizontal(
            OfficeColor.FromRgb(32, 76, 120),
            OfficeColor.FromRgb(78, 159, 188));
        accent.Shadow = new OfficeShadow(OfficeColor.Black, 0.16, 1.5, 1.5);
        accent.StrokeColor = OfficeColor.FromRgb(32, 76, 120);
        accent.StrokeWidth = 0;
        return accent;
    }

    private static PdfTableStyle CreateReportTableStyle() {
        return new PdfTableStyle {
            HeaderFill = PdfColor.FromRgb(32, 76, 120),
            HeaderTextColor = PdfColor.White,
            TextColor = PdfColor.FromRgb(31, 41, 55),
            RowStripeFill = PdfColor.FromRgb(248, 250, 252),
            BorderColor = PdfColor.FromRgb(210, 218, 226),
            BorderWidth = 0.5,
            CellPaddingX = 6,
            CellPaddingY = 5,
            Caption = "Table 1. Report signals",
            CaptionColor = PdfColor.FromRgb(80, 90, 100),
            CaptionFontSize = 8.5,
            CaptionSpacingAfter = 5,
            SpacingBefore = 6,
            AutoFitColumns = true
        };
    }

    private static PdfTableStyle CreateLinksTableStyle() {
        return new PdfTableStyle {
            HeaderFill = PdfColor.FromRgb(32, 76, 120),
            HeaderTextColor = PdfColor.White,
            TextColor = PdfColor.FromRgb(31, 41, 55),
            RowStripeFill = PdfColor.FromRgb(248, 250, 252),
            BorderColor = PdfColor.FromRgb(210, 218, 226),
            BorderWidth = 0.5,
            CellPaddingX = 6,
            CellPaddingY = 5,
            Caption = "Table 1. Linked resources",
            CaptionColor = PdfColor.FromRgb(80, 90, 100),
            CaptionFontSize = 8.5,
            CaptionSpacingAfter = 5,
            SpacingBefore = 6,
            SpacingAfter = 12
        };
    }

    private static void AssertRasterBaseline(string baselineName, string actualPath) {
        string expectedPath = Path.Combine(GetTestsProjectRoot(), "Pdf", "VisualBaselines", baselineName);
        if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_PDF_RASTER_BASELINE"), "1", StringComparison.Ordinal)) {
            Directory.CreateDirectory(Path.GetDirectoryName(expectedPath)!);
            File.Copy(actualPath, expectedPath, overwrite: true);
            return;
        }

        if (!File.Exists(expectedPath)) {
            throw new FileNotFoundException(
                "PDF raster baseline missing. Set OFFICEIMO_UPDATE_PDF_RASTER_BASELINE=1 and re-run this test to generate it.",
                expectedPath);
        }

        RasterComparison comparison = CompareRasterImages(File.ReadAllBytes(expectedPath), File.ReadAllBytes(actualPath));
        if (!comparison.Passed) {
            string artifactDirectory = Path.Combine(Path.GetTempPath(), "OfficeIMO.PdfRaster", DateTime.UtcNow.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(artifactDirectory);

            string actualArtifactPath = Path.Combine(artifactDirectory, Path.GetFileName(actualPath));
            string expectedArtifactPath = Path.Combine(artifactDirectory, Path.GetFileName(expectedPath));
            string diffArtifactPath = Path.Combine(artifactDirectory, Path.GetFileNameWithoutExtension(actualPath) + ".diff.png");
            File.Copy(actualPath, actualArtifactPath, overwrite: true);
            File.Copy(expectedPath, expectedArtifactPath, overwrite: true);
            File.WriteAllBytes(diffArtifactPath, comparison.DiffPng);

            throw new Xunit.Sdk.XunitException(
                "PDF raster baseline changed. " +
                "Different pixels: " + comparison.DifferentPixels + "/" + comparison.TotalPixels + "; " +
                "max channel delta: " + comparison.MaxChannelDelta + "; " +
                "allowed different pixels: " + comparison.AllowedDifferentPixels + "; " +
                "channel tolerance: " + comparison.ChannelTolerance + ". " +
                "Artifacts: " + artifactDirectory + ".");
        }
    }

    private static RasterComparison CompareRasterImages(byte[] expectedPng, byte[] actualPng) {
        int channelTolerance = ReadNonNegativeInt("OFFICEIMO_PDF_RASTER_PIXEL_TOLERANCE", 0);
        int allowedDifferentPixels = ReadNonNegativeInt("OFFICEIMO_PDF_RASTER_ALLOWED_DIFF_PIXELS", 0);
        return CompareRasterImages(expectedPng, actualPng, channelTolerance, allowedDifferentPixels);
    }

    private static RasterComparison CompareRasterImages(byte[] expectedPng, byte[] actualPng, int channelTolerance, int allowedDifferentPixels) {
        PngRaster expected = PngRaster.Decode(expectedPng);
        PngRaster actual = PngRaster.Decode(actualPng);
        if (expected.Width != actual.Width || expected.Height != actual.Height) {
            byte[] sizeDiff = PngRaster.CreateSizeMismatchDiff(expected, actual);
            return new RasterComparison(false, 0, Math.Max(expected.Width * expected.Height, actual.Width * actual.Height), 255, channelTolerance, allowedDifferentPixels, sizeDiff);
        }

        int differentPixels = 0;
        int maxChannelDelta = 0;
        byte[] diff = new byte[expected.Width * expected.Height * 3];

        for (int pixel = 0; pixel < expected.Width * expected.Height; pixel++) {
            int offset = pixel * 4;
            int deltaR = Math.Abs(expected.Pixels[offset] - actual.Pixels[offset]);
            int deltaG = Math.Abs(expected.Pixels[offset + 1] - actual.Pixels[offset + 1]);
            int deltaB = Math.Abs(expected.Pixels[offset + 2] - actual.Pixels[offset + 2]);
            int deltaA = Math.Abs(expected.Pixels[offset + 3] - actual.Pixels[offset + 3]);
            int maxPixelDelta = Math.Max(Math.Max(deltaR, deltaG), Math.Max(deltaB, deltaA));
            maxChannelDelta = Math.Max(maxChannelDelta, maxPixelDelta);

            int diffOffset = pixel * 3;
            if (maxPixelDelta > channelTolerance) {
                differentPixels++;
                diff[diffOffset] = 255;
                diff[diffOffset + 1] = (byte)Math.Min(255, Math.Max(deltaR, deltaG) * 4);
                diff[diffOffset + 2] = (byte)Math.Min(255, Math.Max(deltaB, deltaA) * 4);
            } else {
                int gray = (expected.Pixels[offset] + expected.Pixels[offset + 1] + expected.Pixels[offset + 2]) / 3;
                byte muted = (byte)(240 - Math.Min(120, gray / 3));
                diff[diffOffset] = muted;
                diff[diffOffset + 1] = muted;
                diff[diffOffset + 2] = muted;
            }
        }

        bool passed = differentPixels <= allowedDifferentPixels;
        return new RasterComparison(passed, differentPixels, expected.Width * expected.Height, maxChannelDelta, channelTolerance, allowedDifferentPixels, PngRaster.EncodeRgb(expected.Width, expected.Height, diff));
    }

    private static int ReadNonNegativeInt(string variable, int defaultValue) {
        string? raw = Environment.GetEnvironmentVariable(variable);
        if (string.IsNullOrWhiteSpace(raw)) {
            return defaultValue;
        }

        int value;
        if (!int.TryParse(raw, out value) || value < 0) {
            throw new InvalidOperationException(variable + " must be a non-negative integer.");
        }

        return value;
    }

    private static void RunPdftoppm(string rasterizerPath, string pdfPath, string outputPrefix, string workDir, int pageNumber) {
        string pageText = pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var psi = new ProcessStartInfo {
            FileName = rasterizerPath,
            Arguments = "-r 72 -png -singlefile -f " + pageText + " -l " + pageText + " " + Quote(pdfPath) + " " + Quote(outputPrefix),
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = workDir
        };

        using var process = Process.Start(psi);
        if (process == null) {
            throw new InvalidOperationException("Could not start PDF rasterizer: " + rasterizerPath);
        }

        if (!process.WaitForExit(30000)) {
            process.Kill();
            throw new TimeoutException("PDF rasterizer timed out: " + rasterizerPath);
        }

        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        if (process.ExitCode != 0) {
            throw new InvalidOperationException("PDF rasterizer failed with exit code " + process.ExitCode + "." + Environment.NewLine + output + Environment.NewLine + error);
        }
    }

    private static bool TryFindPdftoppm(out string path) {
        string? configured = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_RASTERIZER");
        if (!string.IsNullOrWhiteSpace(configured) && File.Exists(configured)) {
            path = configured;
            return true;
        }

        string fileName = Environment.OSVersion.Platform == PlatformID.Win32NT ? "pdftoppm.exe" : "pdftoppm";
        string? pathVariable = Environment.GetEnvironmentVariable("PATH");
        if (!string.IsNullOrWhiteSpace(pathVariable)) {
            foreach (string directory in pathVariable.Split(Path.PathSeparator)) {
                if (string.IsNullOrWhiteSpace(directory)) {
                    continue;
                }

                string candidate = Path.Combine(directory.Trim(), fileName);
                if (File.Exists(candidate)) {
                    path = candidate;
                    return true;
                }
            }
        }

        if (Environment.OSVersion.Platform == PlatformID.Win32NT) {
            string? localAppData = Environment.GetEnvironmentVariable("LOCALAPPDATA");
            if (!string.IsNullOrWhiteSpace(localAppData)) {
                string packages = Path.Combine(localAppData, "Microsoft", "WinGet", "Packages");
                if (Directory.Exists(packages)) {
                    string[] candidates = Directory.GetFiles(packages, "pdftoppm.exe", SearchOption.AllDirectories);
                    if (candidates.Length > 0) {
                        path = candidates[0];
                        return true;
                    }
                }
            }
        }

        path = string.Empty;
        return false;
    }

    private static bool IsRequired() =>
        string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_REQUIRE_PDF_RASTERIZER"), "1", StringComparison.Ordinal);

    private static string Quote(string value) => "\"" + value.Replace("\"", "\\\"") + "\"";

    private sealed class RasterComparison {
        internal RasterComparison(bool passed, int differentPixels, int totalPixels, int maxChannelDelta, int channelTolerance, int allowedDifferentPixels, byte[] diffPng) {
            Passed = passed;
            DifferentPixels = differentPixels;
            TotalPixels = totalPixels;
            MaxChannelDelta = maxChannelDelta;
            ChannelTolerance = channelTolerance;
            AllowedDifferentPixels = allowedDifferentPixels;
            DiffPng = diffPng;
        }

        internal bool Passed { get; }
        internal int DifferentPixels { get; }
        internal int TotalPixels { get; }
        internal int MaxChannelDelta { get; }
        internal int ChannelTolerance { get; }
        internal int AllowedDifferentPixels { get; }
        internal byte[] DiffPng { get; }
    }

    private sealed class PngRaster {
        private static readonly byte[] Signature = { 137, 80, 78, 71, 13, 10, 26, 10 };

        private PngRaster(int width, int height, byte[] pixels) {
            Width = width;
            Height = height;
            Pixels = pixels;
        }

        internal int Width { get; }
        internal int Height { get; }
        internal byte[] Pixels { get; }

        internal static PngRaster Decode(byte[] bytes) {
            if (bytes.Length < Signature.Length || !StartsWith(bytes, Signature)) {
                throw new InvalidOperationException("Raster baseline is not a PNG file.");
            }

            int width = 0;
            int height = 0;
            int bitDepth = 0;
            int colorType = 0;
            int compression = 0;
            int filter = 0;
            int interlace = 0;
            var idat = new List<byte>();

            int offset = Signature.Length;
            while (offset + 12 <= bytes.Length) {
                int length = ReadBigEndianInt32(bytes, offset);
                offset += 4;
                string type = Encoding.ASCII.GetString(bytes, offset, 4);
                offset += 4;
                if (length < 0 || offset + length + 4 > bytes.Length) {
                    throw new InvalidOperationException("PNG chunk length is invalid.");
                }

                if (type == "IHDR") {
                    width = ReadBigEndianInt32(bytes, offset);
                    height = ReadBigEndianInt32(bytes, offset + 4);
                    bitDepth = bytes[offset + 8];
                    colorType = bytes[offset + 9];
                    compression = bytes[offset + 10];
                    filter = bytes[offset + 11];
                    interlace = bytes[offset + 12];
                } else if (type == "IDAT") {
                    for (int i = 0; i < length; i++) {
                        idat.Add(bytes[offset + i]);
                    }
                } else if (type == "IEND") {
                    break;
                }

                offset += length + 4;
            }

            if (width <= 0 || height <= 0) {
                throw new InvalidOperationException("PNG image dimensions are invalid.");
            }

            if (bitDepth != 8 || compression != 0 || filter != 0 || interlace != 0 || (colorType != 2 && colorType != 6)) {
                throw new InvalidOperationException("Only non-interlaced 8-bit RGB/RGBA PNG raster baselines are supported.");
            }

            byte[] inflated = InflateZlib(idat.ToArray());
            int channels = colorType == 6 ? 4 : 3;
            int stride = width * channels;
            byte[] pixels = new byte[width * height * 4];
            byte[] previous = new byte[stride];
            byte[] current = new byte[stride];
            int source = 0;
            int bytesPerPixel = channels;

            for (int y = 0; y < height; y++) {
                if (source >= inflated.Length) {
                    throw new InvalidOperationException("PNG image data ended unexpectedly.");
                }

                byte filterType = inflated[source++];
                if (source + stride > inflated.Length) {
                    throw new InvalidOperationException("PNG scanline is incomplete.");
                }

                Buffer.BlockCopy(inflated, source, current, 0, stride);
                source += stride;
                UnfilterScanline(filterType, current, previous, bytesPerPixel);

                for (int x = 0; x < width; x++) {
                    int sourcePixel = x * channels;
                    int targetPixel = (y * width + x) * 4;
                    pixels[targetPixel] = current[sourcePixel];
                    pixels[targetPixel + 1] = current[sourcePixel + 1];
                    pixels[targetPixel + 2] = current[sourcePixel + 2];
                    pixels[targetPixel + 3] = colorType == 6 ? current[sourcePixel + 3] : (byte)255;
                }

                byte[] swap = previous;
                previous = current;
                current = swap;
                Array.Clear(current, 0, current.Length);
            }

            return new PngRaster(width, height, pixels);
        }

        internal static byte[] EncodeRgb(int width, int height, byte[] rgb) {
            if (width <= 0 || height <= 0) {
                throw new ArgumentOutOfRangeException(nameof(width), "PNG dimensions must be positive.");
            }

            if (rgb.Length != width * height * 3) {
                throw new ArgumentException("RGB buffer length does not match PNG dimensions.", nameof(rgb));
            }

            byte[] scanlines = new byte[height * (1 + width * 3)];
            int source = 0;
            int target = 0;
            for (int y = 0; y < height; y++) {
                scanlines[target++] = 0;
                Buffer.BlockCopy(rgb, source, scanlines, target, width * 3);
                source += width * 3;
                target += width * 3;
            }

            using var ms = new MemoryStream();
            ms.Write(Signature, 0, Signature.Length);
            byte[] ihdr = new byte[13];
            WriteBigEndianInt32(ihdr, 0, width);
            WriteBigEndianInt32(ihdr, 4, height);
            ihdr[8] = 8;
            ihdr[9] = 2;
            WriteChunk(ms, "IHDR", ihdr);
            WriteChunk(ms, "IDAT", DeflateZlibStored(scanlines));
            WriteChunk(ms, "IEND", Array.Empty<byte>());
            return ms.ToArray();
        }

        internal static byte[] CreateSizeMismatchDiff(PngRaster expected, PngRaster actual) {
            int width = Math.Max(expected.Width, actual.Width);
            int height = Math.Max(expected.Height, actual.Height);
            byte[] diff = new byte[width * height * 3];
            for (int i = 0; i < diff.Length; i += 3) {
                diff[i] = 255;
                diff[i + 1] = 0;
                diff[i + 2] = 255;
            }

            return EncodeRgb(width, height, diff);
        }

        private static void UnfilterScanline(byte filterType, byte[] current, byte[] previous, int bytesPerPixel) {
            for (int i = 0; i < current.Length; i++) {
                int left = i >= bytesPerPixel ? current[i - bytesPerPixel] : 0;
                int up = previous[i];
                int upLeft = i >= bytesPerPixel ? previous[i - bytesPerPixel] : 0;
                int predictor;
                switch (filterType) {
                    case 0:
                        predictor = 0;
                        break;
                    case 1:
                        predictor = left;
                        break;
                    case 2:
                        predictor = up;
                        break;
                    case 3:
                        predictor = (left + up) / 2;
                        break;
                    case 4:
                        predictor = Paeth(left, up, upLeft);
                        break;
                    default:
                        throw new InvalidOperationException("Unsupported PNG scanline filter: " + filterType + ".");
                }

                current[i] = (byte)((current[i] + predictor) & 0xFF);
            }
        }

        private static int Paeth(int left, int up, int upLeft) {
            int p = left + up - upLeft;
            int pa = Math.Abs(p - left);
            int pb = Math.Abs(p - up);
            int pc = Math.Abs(p - upLeft);
            if (pa <= pb && pa <= pc) return left;
            return pb <= pc ? up : upLeft;
        }

        private static bool StartsWith(byte[] bytes, byte[] prefix) {
            for (int i = 0; i < prefix.Length; i++) {
                if (bytes[i] != prefix[i]) {
                    return false;
                }
            }

            return true;
        }

        private static byte[] InflateZlib(byte[] zlib) {
            if (zlib.Length < 6) {
                throw new InvalidOperationException("PNG zlib stream is too short.");
            }

            using var source = new MemoryStream(zlib, 2, zlib.Length - 6);
            using var deflate = new DeflateStream(source, CompressionMode.Decompress);
            using var output = new MemoryStream();
            deflate.CopyTo(output);
            return output.ToArray();
        }

        private static byte[] DeflateZlibStored(byte[] data) {
            using var ms = new MemoryStream();
            ms.WriteByte(0x78);
            ms.WriteByte(0x01);

            int offset = 0;
            while (offset < data.Length) {
                int blockLength = Math.Min(65535, data.Length - offset);
                bool final = offset + blockLength >= data.Length;
                ms.WriteByte(final ? (byte)1 : (byte)0);
                ms.WriteByte((byte)(blockLength & 0xFF));
                ms.WriteByte((byte)((blockLength >> 8) & 0xFF));
                int nlen = blockLength ^ 0xFFFF;
                ms.WriteByte((byte)(nlen & 0xFF));
                ms.WriteByte((byte)((nlen >> 8) & 0xFF));
                ms.Write(data, offset, blockLength);
                offset += blockLength;
            }

            uint adler = Adler32(data);
            ms.WriteByte((byte)((adler >> 24) & 0xFF));
            ms.WriteByte((byte)((adler >> 16) & 0xFF));
            ms.WriteByte((byte)((adler >> 8) & 0xFF));
            ms.WriteByte((byte)(adler & 0xFF));
            return ms.ToArray();
        }

        private static uint Adler32(byte[] data) {
            const uint mod = 65521;
            uint a = 1;
            uint b = 0;
            for (int i = 0; i < data.Length; i++) {
                a = (a + data[i]) % mod;
                b = (b + a) % mod;
            }

            return (b << 16) | a;
        }

        private static int ReadBigEndianInt32(byte[] bytes, int offset) =>
            (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

        private static void WriteBigEndianInt32(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)((value >> 24) & 0xFF);
            bytes[offset + 1] = (byte)((value >> 16) & 0xFF);
            bytes[offset + 2] = (byte)((value >> 8) & 0xFF);
            bytes[offset + 3] = (byte)(value & 0xFF);
        }

        private static void WriteChunk(Stream stream, string type, byte[] data) {
            byte[] typeBytes = Encoding.ASCII.GetBytes(type);
            byte[] length = new byte[4];
            WriteBigEndianInt32(length, 0, data.Length);
            stream.Write(length, 0, length.Length);
            stream.Write(typeBytes, 0, typeBytes.Length);
            stream.Write(data, 0, data.Length);

            uint crc = Crc32(typeBytes, data);
            byte[] crcBytes = new byte[4];
            WriteBigEndianInt32(crcBytes, 0, unchecked((int)crc));
            stream.Write(crcBytes, 0, crcBytes.Length);
        }

        private static uint Crc32(byte[] type, byte[] data) {
            uint crc = 0xFFFFFFFF;
            for (int i = 0; i < type.Length; i++) {
                crc = UpdateCrc(crc, type[i]);
            }

            for (int i = 0; i < data.Length; i++) {
                crc = UpdateCrc(crc, data[i]);
            }

            return crc ^ 0xFFFFFFFF;
        }

        private static uint UpdateCrc(uint crc, byte value) {
            crc ^= value;
            for (int i = 0; i < 8; i++) {
                crc = (crc & 1) != 0 ? 0xEDB88320 ^ (crc >> 1) : crc >> 1;
            }

            return crc;
        }
    }

    private static string GetTestsProjectRoot() {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.Tests.csproj"))) {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate OfficeIMO.Tests project root from test runtime base directory.");
    }

    private static void TryDeleteDirectory(string directory) {
        try {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        } catch {
        }
    }

    private static byte[] CreateFallbackLogo() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }

    private static byte[] CreateTransparentBadgePng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 6, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 16,
            73, 68, 65, 84,
            0x78, 0x01, 0x01, 0x05, 0x00, 0xFA, 0xFF, 0x00,
            0x2A, 0x84, 0x52, 0xA0, 0x03, 0x7D, 0x01, 0xA1,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }
}
