using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentRasterVisualBaselineTests {
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
        options.UseOfficeFontFamily("Arial");

        return PdfDocument.Create(options)
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
        return CreateVisualBaselineDocument(new PdfOptions {
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
            }.UseOfficeFontFamily("Arial"))
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
        return CreateVisualBaselineDocument(new PdfOptions {
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
            }.UseOfficeFontFamily("Arial"))
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

    private static byte[] CreateWatermark() {
        return CreateVisualBaselineDocument(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf watermark gate",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Watermark", author: "OfficeIMO")
            .Watermark("DRAFT", fontSize: 74, color: PdfColor.FromRgb(90, 106, 130), opacity: 0.14, rotationAngle: -38)
            .H1("Watermark Gate", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("This baseline verifies reusable page watermarks render behind normal document content without taking layout space or reducing text readability."))
            .PanelParagraph(p => p
                    .Bold("Document state").LineBreak()
                    .Text("Watermark text is a page decoration, not a body block, so paragraphs, tables, and footers keep their normal rhythm."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(248, 250, 252),
                    BorderColor = PdfColor.FromRgb(203, 213, 225),
                    BorderWidth = 0.7,
                    PaddingX = 10,
                    PaddingY = 8,
                    SpacingBefore = 8,
                    SpacingAfter = 10
                })
            .Table(new[] {
                new[] { "Capability", "Visual expectation" },
                new[] { "Opacity", "Subtle behind-content mark" },
                new[] { "Rotation", "Centered diagonal placement" },
                new[] { "Flow", "No extra body spacing" }
            }, style: new PdfTableStyle {
                HeaderFill = PdfColor.FromRgb(15, 23, 42),
                HeaderTextColor = PdfColor.White,
                BorderColor = PdfColor.FromRgb(203, 213, 225),
                RowStripeFill = PdfColor.FromRgb(248, 250, 252),
                CellPaddingX = 7,
                CellPaddingY = 5,
                AutoFitColumns = true,
                SpacingBefore = 4,
                SpacingAfter = 8
            })
            .Paragraph(p => p.Text("The watermark remains readable in the raster gate while the foreground table and panel stay crisp."))
            .ToBytes();
    }

    private static byte[] CreateImageWatermark() {
        string logoPath = Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png");
        byte[] logo = File.Exists(logoPath) ? File.ReadAllBytes(logoPath) : CreateFallbackLogo();

        return CreateVisualBaselineDocument(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf image watermark gate",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Image Watermark", author: "OfficeIMO")
            .ImageWatermark(logo, width: 220, height: 92, opacity: 0.16, rotationAngle: -22)
            .H1("Image Watermark Gate", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("This baseline verifies reusable image watermarks render behind normal document content without occupying flow space."))
            .PanelParagraph(p => p
                    .Bold("Layering expectation").LineBreak()
                    .Text("The image watermark is a page decoration. Foreground text, panels, and table borders stay readable and crisp."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(248, 250, 252),
                    BorderColor = PdfColor.FromRgb(203, 213, 225),
                    BorderWidth = 0.7,
                    PaddingX = 10,
                    PaddingY = 8,
                    SpacingBefore = 8,
                    SpacingAfter = 10
                })
            .Table(new[] {
                new[] { "Capability", "Visual expectation" },
                new[] { "Image XObject", "Drawn behind body content" },
                new[] { "Opacity", "Subtle enough for text readability" },
                new[] { "Rotation", "Centered page decoration" }
            }, style: new PdfTableStyle {
                HeaderFill = PdfColor.FromRgb(15, 23, 42),
                HeaderTextColor = PdfColor.White,
                BorderColor = PdfColor.FromRgb(203, 213, 225),
                RowStripeFill = PdfColor.FromRgb(248, 250, 252),
                CellPaddingX = 7,
                CellPaddingY = 5,
                AutoFitColumns = true,
                SpacingBefore = 4,
                SpacingAfter = 8
            })
            .Paragraph(p => p.Text("The raster gate keeps this from quietly regressing into an above-content stamp or layout-consuming image block."))
            .ToBytes();
    }

    private static byte[] CreatePageBorder() {
        return CreateVisualBaselineDocument(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf page border gate",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Page Border", author: "OfficeIMO")
            .Background(PdfColor.FromRgb(250, 252, 255))
            .PageBorder(PdfColor.FromRgb(30, 64, 175), width: 1.15, inset: 30, opacity: 0.72, dashStyle: OfficeStrokeDashStyle.Solid)
            .H1("Page Border Gate", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("This baseline verifies reusable page borders render as decoration without taking layout space or crowding text."))
            .PanelParagraph(p => p
                    .Bold("Frame expectation").LineBreak()
                    .Text("The frame should visually finish the page while content keeps normal margins, table rhythm, and footer placement."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(255, 255, 255),
                    BorderColor = PdfColor.FromRgb(191, 219, 254),
                    BorderWidth = 0.7,
                    PaddingX = 10,
                    PaddingY = 8,
                    SpacingBefore = 8,
                    SpacingAfter = 10
                })
            .Table(new[] {
                new[] { "Capability", "Visual expectation" },
                new[] { "Inset", "Frame stays clear of page edges" },
                new[] { "Opacity", "Border is visible without dominating" },
                new[] { "Flow", "No extra body spacing" }
            }, style: new PdfTableStyle {
                HeaderFill = PdfColor.FromRgb(15, 23, 42),
                HeaderTextColor = PdfColor.White,
                BorderColor = PdfColor.FromRgb(203, 213, 225),
                RowStripeFill = PdfColor.FromRgb(248, 250, 252),
                CellPaddingX = 7,
                CellPaddingY = 5,
                AutoFitColumns = true,
                SpacingBefore = 4,
                SpacingAfter = 8
            })
            .Paragraph(p => p.Text("The border is a generic page primitive that Word, Markdown, Excel, and wrappers can reuse without template-specific APIs."))
            .ToBytes();
    }

    private static byte[] CreateBackgroundImage() {
        string logoPath = Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png");
        byte[] logo = File.Exists(logoPath) ? File.ReadAllBytes(logoPath) : CreateFallbackLogo();

        return CreateVisualBaselineDocument(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf background image gate",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Background Image", author: "OfficeIMO")
            .Background(PdfColor.FromRgb(250, 252, 255))
            .BackgroundImage(logo, OfficeImageFit.Contain, opacity: 0.08)
            .PageBorder(PdfColor.FromRgb(148, 163, 184), width: 0.9, inset: 30, opacity: 0.55)
            .H1("Background Image Gate", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("This baseline verifies a page background image can fill the visual surface without becoming a body image."))
            .PanelParagraph(p => p
                    .Bold("Layering expectation").LineBreak()
                    .Text("The background image is fitted to the page box behind content, watermarks, borders, headers, footers, and normal flow blocks."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(255, 255, 255),
                    BorderColor = PdfColor.FromRgb(203, 213, 225),
                    BorderWidth = 0.7,
                    PaddingX = 10,
                    PaddingY = 8,
                    SpacingBefore = 8,
                    SpacingAfter = 10
                })
            .Table(new[] {
                new[] { "Capability", "Visual expectation" },
                new[] { "Fit", "Contain/cover/stretch without layout space" },
                new[] { "Opacity", "Subtle letterhead-style background" },
                new[] { "Flow", "Foreground rhythm remains stable" }
            }, style: new PdfTableStyle {
                HeaderFill = PdfColor.FromRgb(15, 23, 42),
                HeaderTextColor = PdfColor.White,
                BorderColor = PdfColor.FromRgb(203, 213, 225),
                RowStripeFill = PdfColor.FromRgb(248, 250, 252),
                CellPaddingX = 7,
                CellPaddingY = 5,
                AutoFitColumns = true,
                SpacingBefore = 4,
                SpacingAfter = 8
            })
            .Paragraph(p => p.Text("The same primitive can later power Markdown themes, Word backgrounds, letterheads, and wrapper presets."))
            .ToBytes();
    }

    private static byte[] CreateBackgroundShapes() {
        return CreateVisualBaselineDocument(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf background shapes gate",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            })
            .Meta(title: "OfficeIMO.Pdf Background Shapes", author: "OfficeIMO")
            .Background(PdfColor.FromRgb(250, 252, 255))
            .BackgroundTopBand(104, insetX: 36, offsetY: 58, cornerRadius: 22, stroke: PdfColor.FromRgb(147, 197, 253), strokeWidth: 0.7, fillOpacity: 0.64, strokeOpacity: 0.6, fillGradient: OfficeLinearGradient.Horizontal(OfficeColor.FromRgb(219, 234, 254), OfficeColor.FromRgb(240, 253, 250)))
            .BackgroundRightBand(70, PdfColor.FromRgb(239, 246, 255), insetY: 92, offsetX: 30, cornerRadius: 32, fillOpacity: 0.6)
            .BackgroundEllipse(456, 92, 190, 190, PdfColor.FromRgb(224, 242, 254), fillOpacity: 0.45)
            .PageBorder(PdfColor.FromRgb(148, 163, 184), width: 0.9, inset: 30, opacity: 0.55)
            .H1("Background Shapes Gate", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("This baseline verifies vector page decoration can add Word-like polish without taking body layout space."))
            .PanelParagraph(p => p
                    .Bold("Layering expectation").LineBreak()
                    .Text("Rounded bands, ellipses, gradients, opacity, borders, headers, footers, and foreground flow should stay in a stable visual stack."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(255, 255, 255),
                    BorderColor = PdfColor.FromRgb(203, 213, 225),
                    BorderWidth = 0.7,
                    PaddingX = 10,
                    PaddingY = 8,
                    SpacingBefore = 8,
                    SpacingAfter = 10
                })
            .Table(new[] {
                new[] { "Capability", "Visual expectation" },
                new[] { "Vector fills", "No raster blur, reusable page decoration" },
                new[] { "Opacity", "Soft bands behind normal text" },
                new[] { "Flow", "Content rhythm remains stable" }
            }, style: new PdfTableStyle {
                HeaderFill = PdfColor.FromRgb(15, 23, 42),
                HeaderTextColor = PdfColor.White,
                BorderColor = PdfColor.FromRgb(203, 213, 225),
                RowStripeFill = PdfColor.FromRgb(248, 250, 252),
                CellPaddingX = 7,
                CellPaddingY = 5,
                AutoFitColumns = true,
                SpacingBefore = 4,
                SpacingAfter = 8
            })
            .Paragraph(p => p.Text("This gives Markdown and Office exporters a dependency-free way to create visually richer pages than literal plain Markdown output."))
            .ToBytes();
    }

    private static byte[] CreateRowColumns() {
        return CreateVisualBaselineDocument(new PdfOptions {
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
        return CreateVisualBaselineDocument(new PdfOptions {
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
        return CreateVisualBaselineDocument(new PdfOptions {
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
            new[] { "Flow", "Generic", "Borders and row separators should reveal each preset shape at raster level." }
        };

        PdfDocument doc = CreateVisualBaselineDocument(new PdfOptions {
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

        string[] visualStyleNames = {
            "TableGrid",
            "TableGridLight",
            "PlainTable1",
            "GridTable1Light",
            "ListTable1Light"
        };

        foreach (string styleName in visualStyleNames) {
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

        doc.Table(CreateWordAccentSwatchRows(), PdfAlign.Left, CreateWordAccentSwatchStyle());

        return doc.ToBytes();
    }

    private static string[][] CreateWordAccentSwatchRows() {
        return new[] {
            new[] { "Role", "A1", "A2", "A3", "A4", "A5", "A6" },
            new[] { "Grid line", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty },
            new[] { "Strong line", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty },
            new[] { "List band", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty }
        };
    }

    private static PdfTableStyle CreateWordAccentSwatchStyle() {
        var style = TableStyles.ListTable1Light();
        style.Caption = "Accent swatches";
        style.CaptionColor = PdfColor.FromRgb(80, 90, 100);
        style.CaptionFontSize = 8.5;
        style.CaptionSpacingAfter = 4;
        style.SpacingBefore = 6;
        style.SpacingAfter = 4;
        style.FontSize = 8.5;
        style.HeaderFontSize = 8.5;
        style.CellPaddingX = 5;
        style.CellPaddingY = 3;
        style.ColumnWidthPoints = new List<double?> { 76, 45, 45, 45, 45, 45, 45 };
        style.AutoFitColumns = false;
        style.Alignments = new List<PdfColumnAlign> {
            PdfColumnAlign.Left,
            PdfColumnAlign.Center,
            PdfColumnAlign.Center,
            PdfColumnAlign.Center,
            PdfColumnAlign.Center,
            PdfColumnAlign.Center,
            PdfColumnAlign.Center
        };

        var cellFills = new Dictionary<(int Row, int Column), PdfColor>();
        for (int accent = 1; accent <= 6; accent++) {
            PdfTableStyle grid = TableStyles.FromWordTableStyle("GridTable1LightAccent" + accent.ToString(System.Globalization.CultureInfo.InvariantCulture));
            PdfTableStyle list = TableStyles.FromWordTableStyle("ListTable1LightAccent" + accent.ToString(System.Globalization.CultureInfo.InvariantCulture));
            cellFills[(1, accent)] = grid.BorderColor ?? PdfColor.White;
            cellFills[(2, accent)] = grid.HeaderSeparatorColor ?? PdfColor.White;
            cellFills[(3, accent)] = list.RowStripeFill ?? PdfColor.White;
        }

        style.CellFills = cellFills;
        return style;
    }
}
