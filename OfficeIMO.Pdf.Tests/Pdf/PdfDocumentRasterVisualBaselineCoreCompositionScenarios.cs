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
        var options = new PdfCreateOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
            DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
            PageSetup = {
                Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8) },
                Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8) }
            }
        };
        return PdfDocument.Create(options)
            .Compose(document => {
                document.Page(page => {
                    page.Size(PdfPageSizes.A5);
                    page.Margin(global::OfficeIMO.OfficeLength.Points(36), global::OfficeIMO.OfficeLength.Points(42), global::OfficeIMO.OfficeLength.Points(36), global::OfficeIMO.OfficeLength.Points(42));
                    page.DefaultTextStyle(x => x
                        .Font(PdfStandardFont.Helvetica)
                        .FontSize(global::OfficeIMO.OfficeLength.Points(10))
                        .Color(global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55)));
                    page.Header(h => h.Alignment(OfficeAlignment.Left).Text("OfficeIMO.Pdf compose DSL gate"));

                    page.Content(c => c
                        .Column(column => {
                            column.Item().H1("Compose DSL");
                            column.Item().AddParagraph(p => p.AddText("A compact multi-page visual baseline for the OfficeIMO.Pdf composition API."));
                            column.Item().AddPanelParagraph(
                                p => {
                                    p.AddText("What this protects").SetBold();
                                    p.AddLineBreak();
                                    p.AddText("Page settings, composed content, explicit page breaks, header/footer tokens, and rich text inside composed items.");
                                },
                                new PdfPanelStyle {
                                    Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                                    PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                                    PaddingY = global::OfficeIMO.OfficeLength.Points(7)
                                });
                            column.Item().PageBreak();

                            column.Item().H2("Operational Notes");
                            column.Item().AddList(new[] {
                                "Compose uses the same document engine as fluent blocks.",
                                "Visual gates should cover both public authoring styles.",
                                "Footer page totals must remain correct after explicit breaks."
                            }, color: global::OfficeIMO.OfficeColor.FromRgb(55, 65, 81));
                            column.Item().HR(global::OfficeIMO.OfficeLength.Points(0.8), global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207), global::OfficeIMO.OfficeLength.Points(8), global::OfficeIMO.OfficeLength.Points(8));
                            column.Item().AddParagraph(p => {
                                p.AddText("Status: ");
                                p.AddText("ready for wrapper experiments ").SetBold();
                                p.AddText("once visual quality keeps improving.").SetColor(global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180));
                            });
                            column.Item().PageBreak();

                            column.Item().H2("Color Sections");
                            column.Item().AddParagraph(p => p.AddText("Critical items should remain readable when colored inline.").SetColor(global::OfficeIMO.OfficeColor.FromRgb(185, 28, 28)));
                            column.Item().AddParagraph(p => p.AddText("Healthy items should use calm green without overpowering the page.").SetColor(global::OfficeIMO.OfficeColor.FromRgb(22, 101, 52)));
                            column.Item().AddParagraph(p => p.AddText("Informational items should keep good contrast on white backgrounds.").SetColor(global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180)));
                            column.Item().AddParagraph(p => p.AddText("End of compose DSL sample."), OfficeAlignment.Right, global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80));
                        }));

                    page.Footer(f => f.Alignment(OfficeAlignment.Center).Text(t => t
                        .Text("OfficeIMO.Pdf compose - page ")
                        .CurrentPage()
                        .Text("/")
                        .TotalPages()));
                });
            })
            .ToBytes();
    }

    internal static byte[] CreateShowcaseDashboard() {
        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9.5),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(30, 41, 59),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf showcase - landscape dashboard", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                },
                CreateOutlineFromHeadings = true
            })
            .Meta(
                title: "OfficeIMO.Pdf Showcase Dashboard",
                author: "OfficeIMO",
                subject: "Landscape dashboard from generic PDF primitives",
                keywords: "OfficeIMO,Pdf,dashboard,drawings,tables")
            .Compose(document => {
                document.Page(page => {
                    page.Size(PdfPageSizes.A4).Landscape().Margin(global::OfficeIMO.OfficeLength.Points(42), global::OfficeIMO.OfficeLength.Points(38), global::OfficeIMO.OfficeLength.Points(42), global::OfficeIMO.OfficeLength.Points(42));
                    page.DefaultParagraphStyle(new PdfParagraphStyle { LineHeight = 1.18, SpacingAfter = global::OfficeIMO.OfficeLength.Points(5) });
                    page.Content(content => {
                        content.Item(item => item.H1("Quarterly Operations Dashboard", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42)));
                        content.Item(item => item.AddParagraph(p => p.AddText("A single-page control surface composed from rows, panels, reusable drawings, wrapped tables, and compact list rhythm."), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(10) }));

                        content.Row(row => {
                            row.Gap(global::OfficeIMO.OfficeLength.Points(14))
                                .PercentColumn(25, column => column.AddPanelParagraph(p => { p.AddText("92%").SetBold(); p.AddText("\nSLA attainment"); }, CreateDashboardMetricPanel(global::OfficeIMO.OfficeColor.FromRgb(236, 253, 245), global::OfficeIMO.OfficeColor.FromRgb(22, 163, 74)), OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(22, 101, 52)))
                                .PercentColumn(25, column => column.AddPanelParagraph(p => { p.AddText("1.8h").SetBold(); p.AddText("\nMean response"); }, CreateDashboardMetricPanel(global::OfficeIMO.OfficeColor.FromRgb(239, 246, 255), global::OfficeIMO.OfficeColor.FromRgb(37, 99, 235)), OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(30, 64, 175)))
                                .PercentColumn(25, column => column.AddPanelParagraph(p => { p.AddText("34").SetBold(); p.AddText("\nOpen actions"); }, CreateDashboardMetricPanel(global::OfficeIMO.OfficeColor.FromRgb(255, 251, 235), global::OfficeIMO.OfficeColor.FromRgb(217, 119, 6)), OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(146, 64, 14)))
                                .PercentColumn(25, column => column.AddPanelParagraph(p => { p.AddText("0").SetBold(); p.AddText("\nCritical blockers"); }, CreateDashboardMetricPanel(global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252), global::OfficeIMO.OfficeColor.FromRgb(100, 116, 139)), OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(51, 65, 85)));
                        });

                        content.Spacer(global::OfficeIMO.OfficeLength.Points(12));
                        content.Row(row => {
                            row.Gap(global::OfficeIMO.OfficeLength.Points(18))
                                .PercentColumn(58, column => {
                                    column.AddParagraph(p => p.AddText("Delivery trend").SetBold(), style: CreateDashboardSectionLabelStyle());
                                    column.Drawing(CreateDashboardTrendDrawing(), OfficeAlignment.Left, spacingBefore: global::OfficeIMO.OfficeLength.Points(2), spacingAfter: global::OfficeIMO.OfficeLength.Points(8));
                                    column.AddTable(CreateDashboardRiskRows(), style: CreateDashboardRiskTableStyle());
                                })
                                .PercentColumn(42, column => {
                                    column.AddParagraph(p => p.AddText("Narrative").SetBold(), style: CreateDashboardSectionLabelStyle());
                                    column.AddPanelParagraph(
                                        p => p.AddText("The dashboard deliberately avoids a domain-specific report object. It uses the same primitive surface that a Word, Excel, or PowerPoint exporter could target later: page setup, rows, tables, paragraphs, shapes, and themes."),
                                        new PdfPanelStyle {
                                            Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                                            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225),
                                            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.7),
                                            PaddingX = global::OfficeIMO.OfficeLength.Points(10),
                                            PaddingY = global::OfficeIMO.OfficeLength.Points(8)
                                        });
                                    column.AddList(new[] {
                                        "Rows keep gutters as layout state.",
                                        "Tables use explicit widths and numeric alignment.",
                                        "Vector drawing comes from OfficeIMO.Drawing descriptors.",
                                        "Visual gates can rasterize the result and catch rhythm regressions."
                                    }, style: new PdfListStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(4), ItemSpacing = global::OfficeIMO.OfficeLength.Points(2) });
                                    column.AddTable(CreateDashboardDecisionRows(), style: CreateDashboardDecisionTableStyle());
                                });
                        });
                    });
                });
            })
            .ToBytes();
    }

    private static PdfPanelStyle CreateDashboardMetricPanel(global::OfficeIMO.OfficeColor background, global::OfficeIMO.OfficeColor border) {
        return new PdfPanelStyle {
            Background = background,
            BorderColor = border,
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.8),
            PaddingX = global::OfficeIMO.OfficeLength.Points(10),
            PaddingY = global::OfficeIMO.OfficeLength.Points(8)
        };
    }

    private static PdfParagraphStyle CreateDashboardSectionLabelStyle() {
        return new PdfParagraphStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(4), LineHeight = 1.1 };
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
        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Drawing shared vector gate", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Drawing Gallery", author: "OfficeIMO")
            .H1("Drawing Gallery", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("A visual baseline for shared OfficeIMO.Drawing vector descriptors rendered by the dependency-free PDF engine."))
            .Drawing(CreateDrawingScene(), OfficeAlignment.Center, spacingBefore: 8, spacingAfter: 10)
            .Paragraph(p => {
                p.AddText("Covered: ");
                p.AddText("gradients").SetBold();
                p.AddText(", shadows, dashed strokes, line caps and joins, clipping paths, transforms, grouped scenes, and freeform paths.");
            })
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
        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf watermark gate", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Watermark", author: "OfficeIMO")
            .Watermark("DRAFT", fontSize: 74, color: global::OfficeIMO.OfficeColor.FromRgb(90, 106, 130), opacity: 0.14, rotationAngle: -38)
            .H1("Watermark Gate", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("This baseline verifies reusable page watermarks render behind normal document content without taking layout space or reducing text readability."))
            .PanelParagraph(p => {
                p.AddText("Document state").SetBold();
                p.AddLineBreak();
                p.AddText("Watermark text is a page decoration, not a body block, so paragraphs, tables, and footers keep their normal rhythm.");
            },
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225),
                    BorderWidth = global::OfficeIMO.OfficeLength.Points(0.7),
                    PaddingX = global::OfficeIMO.OfficeLength.Points(10),
                    PaddingY = global::OfficeIMO.OfficeLength.Points(8),
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(8),
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(10)
                })
            .Table(new[] {
                new[] { "Capability", "Visual expectation" },
                new[] { "Opacity", "Subtle behind-content mark" },
                new[] { "Rotation", "Centered diagonal placement" },
                new[] { "Flow", "No extra body spacing" }
            }, style: new PdfTableStyle {
                HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42),
                HeaderTextColor = global::OfficeIMO.OfficeColor.White,
                BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225),
                RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(7),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(5),
                AutoFitColumns = true,
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(4),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(8)
            })
            .Paragraph(p => p.AddText("The watermark remains readable in the raster gate while the foreground table and panel stay crisp."))
            .ToBytes();
    }

    private static byte[] CreateImageWatermark() {
        string logoPath = Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png");
        byte[] logo = File.Exists(logoPath) ? File.ReadAllBytes(logoPath) : CreateFallbackLogo();

        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf image watermark gate", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Image Watermark", author: "OfficeIMO")
            .ImageWatermark(logo, width: 220, height: 92, opacity: 0.16, rotationAngle: -22)
            .H1("Image Watermark Gate", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("This baseline verifies reusable image watermarks render behind normal document content without occupying flow space."))
            .PanelParagraph(p => {
                p.AddText("Layering expectation").SetBold();
                p.AddLineBreak();
                p.AddText("The image watermark is a page decoration. Foreground text, panels, and table borders stay readable and crisp.");
            },
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225),
                    BorderWidth = global::OfficeIMO.OfficeLength.Points(0.7),
                    PaddingX = global::OfficeIMO.OfficeLength.Points(10),
                    PaddingY = global::OfficeIMO.OfficeLength.Points(8),
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(8),
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(10)
                })
            .Table(new[] {
                new[] { "Capability", "Visual expectation" },
                new[] { "Image XObject", "Drawn behind body content" },
                new[] { "Opacity", "Subtle enough for text readability" },
                new[] { "Rotation", "Centered page decoration" }
            }, style: new PdfTableStyle {
                HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42),
                HeaderTextColor = global::OfficeIMO.OfficeColor.White,
                BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225),
                RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(7),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(5),
                AutoFitColumns = true,
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(4),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(8)
            })
            .Paragraph(p => p.AddText("The raster gate keeps this from quietly regressing into an above-content stamp or layout-consuming image block."))
            .ToBytes();
    }

    private static byte[] CreatePageBorder() {
        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf page border gate", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Page Border", author: "OfficeIMO")
            .Background(global::OfficeIMO.OfficeColor.FromRgb(250, 252, 255))
            .PageBorder(global::OfficeIMO.OfficeColor.FromRgb(30, 64, 175), width: 1.15, inset: 30, opacity: 0.72, dashStyle: OfficeStrokeDashStyle.Solid)
            .H1("Page Border Gate", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("This baseline verifies reusable page borders render as decoration without taking layout space or crowding text."))
            .PanelParagraph(p => {
                p.AddText("Frame expectation").SetBold();
                p.AddLineBreak();
                p.AddText("The frame should visually finish the page while content keeps normal margins, table rhythm, and footer placement.");
            },
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(255, 255, 255),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(191, 219, 254),
                    BorderWidth = global::OfficeIMO.OfficeLength.Points(0.7),
                    PaddingX = global::OfficeIMO.OfficeLength.Points(10),
                    PaddingY = global::OfficeIMO.OfficeLength.Points(8),
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(8),
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(10)
                })
            .Table(new[] {
                new[] { "Capability", "Visual expectation" },
                new[] { "Inset", "Frame stays clear of page edges" },
                new[] { "Opacity", "Border is visible without dominating" },
                new[] { "Flow", "No extra body spacing" }
            }, style: new PdfTableStyle {
                HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42),
                HeaderTextColor = global::OfficeIMO.OfficeColor.White,
                BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225),
                RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(7),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(5),
                AutoFitColumns = true,
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(4),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(8)
            })
            .Paragraph(p => p.AddText("The border is a generic page primitive that Word, Markdown, Excel, and wrappers can reuse without template-specific APIs."))
            .ToBytes();
    }

    private static byte[] CreateBackgroundImage() {
        string logoPath = Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png");
        byte[] logo = File.Exists(logoPath) ? File.ReadAllBytes(logoPath) : CreateFallbackLogo();

        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf background image gate", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Background Image", author: "OfficeIMO")
            .Background(global::OfficeIMO.OfficeColor.FromRgb(250, 252, 255))
            .BackgroundImage(logo, OfficeImageFit.Contain, opacity: 0.08)
            .PageBorder(global::OfficeIMO.OfficeColor.FromRgb(148, 163, 184), width: 0.9, inset: 30, opacity: 0.55)
            .H1("Background Image Gate", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("This baseline verifies a page background image can fill the visual surface without becoming a body image."))
            .PanelParagraph(p => {
                p.AddText("Layering expectation").SetBold();
                p.AddLineBreak();
                p.AddText("The background image is fitted to the page box behind content, watermarks, borders, headers, footers, and normal flow blocks.");
            },
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(255, 255, 255),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225),
                    BorderWidth = global::OfficeIMO.OfficeLength.Points(0.7),
                    PaddingX = global::OfficeIMO.OfficeLength.Points(10),
                    PaddingY = global::OfficeIMO.OfficeLength.Points(8),
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(8),
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(10)
                })
            .Table(new[] {
                new[] { "Capability", "Visual expectation" },
                new[] { "Fit", "Contain/cover/stretch without layout space" },
                new[] { "Opacity", "Subtle letterhead-style background" },
                new[] { "Flow", "Foreground rhythm remains stable" }
            }, style: new PdfTableStyle {
                HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42),
                HeaderTextColor = global::OfficeIMO.OfficeColor.White,
                BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225),
                RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(7),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(5),
                AutoFitColumns = true,
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(4),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(8)
            })
            .Paragraph(p => p.AddText("The same primitive can later power Markdown themes, Word backgrounds, letterheads, and wrapper presets."))
            .ToBytes();
    }

    private static byte[] CreateBackgroundShapes() {
        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf background shapes gate", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Background Shapes", author: "OfficeIMO")
            .Background(global::OfficeIMO.OfficeColor.FromRgb(250, 252, 255))
            .BackgroundTopBand(104, insetX: 36, offsetY: 58, cornerRadius: 22, stroke: global::OfficeIMO.OfficeColor.FromRgb(147, 197, 253), strokeWidth: 0.7, fillOpacity: 0.64, strokeOpacity: 0.6, fillGradient: OfficeLinearGradient.Horizontal(OfficeColor.FromRgb(219, 234, 254), OfficeColor.FromRgb(240, 253, 250)))
            .BackgroundRightBand(70, global::OfficeIMO.OfficeColor.FromRgb(239, 246, 255), insetY: 92, offsetX: 30, cornerRadius: 32, fillOpacity: 0.6)
            .BackgroundEllipse(456, 92, 190, 190, global::OfficeIMO.OfficeColor.FromRgb(224, 242, 254), fillOpacity: 0.45)
            .PageBorder(global::OfficeIMO.OfficeColor.FromRgb(148, 163, 184), width: 0.9, inset: 30, opacity: 0.55)
            .H1("Background Shapes Gate", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("This baseline verifies vector page decoration can add Word-like polish without taking body layout space."))
            .PanelParagraph(p => {
                p.AddText("Layering expectation").SetBold();
                p.AddLineBreak();
                p.AddText("Rounded bands, ellipses, gradients, opacity, borders, headers, footers, and foreground flow should stay in a stable visual stack.");
            },
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(255, 255, 255),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225),
                    BorderWidth = global::OfficeIMO.OfficeLength.Points(0.7),
                    PaddingX = global::OfficeIMO.OfficeLength.Points(10),
                    PaddingY = global::OfficeIMO.OfficeLength.Points(8),
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(8),
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(10)
                })
            .Table(new[] {
                new[] { "Capability", "Visual expectation" },
                new[] { "Vector fills", "No raster blur, reusable page decoration" },
                new[] { "Opacity", "Soft bands behind normal text" },
                new[] { "Flow", "Content rhythm remains stable" }
            }, style: new PdfTableStyle {
                HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42),
                HeaderTextColor = global::OfficeIMO.OfficeColor.White,
                BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225),
                RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(7),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(5),
                AutoFitColumns = true,
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(4),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(8)
            })
            .Paragraph(p => p.AddText("This gives Markdown and Office exporters a dependency-free way to create visually richer pages than literal plain Markdown output."))
            .ToBytes();
    }

    private static byte[] CreateRowColumns() {
        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf row columns", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Row Columns", author: "OfficeIMO")
            .Compose(document => {
                document.Page(page => {
                    page.Content(content => {
                        content.Column(column => {
                            column.Item().H1("Row Columns");
                            column.Item().AddParagraph(p => p.AddText("A compact visual gate for composed columns with first-class gutters and independent column flow."));
                        });
                        content.Row(row => {
                            row.Gap(global::OfficeIMO.OfficeLength.Points(18));
                            row.PercentColumn(50, column => column
                                .H2("Status")
                                .AddParagraph(p => p.AddText("The left column carries operational copy with comfortable wrapping, spacing, and no collision with the neighboring column."))
                                .AddList(new[] {
                                    "Gutters are first-class layout state.",
                                    "List markers stay inside the column frame."
                                }, color: global::OfficeIMO.OfficeColor.FromRgb(55, 65, 81))
                                .AddPanelParagraph(
                                    p => { p.AddText("Callout: ").SetBold(); p.AddText("column panels can hold emphasis without leaving the row flow."); },
                                    new PdfPanelStyle {
                                        Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                                        BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                                        PaddingX = global::OfficeIMO.OfficeLength.Points(7),
                                        PaddingY = global::OfficeIMO.OfficeLength.Points(5),
                                        KeepTogether = true
                                    })
                                .RoundedRectangle(global::OfficeIMO.OfficeLength.Points(96), global::OfficeIMO.OfficeLength.Points(5), global::OfficeIMO.OfficeLength.Points(2.5), strokeColor: global::OfficeIMO.OfficeColor.FromRgb(22, 101, 52), strokeWidth: global::OfficeIMO.OfficeLength.Points(0), fillColor: global::OfficeIMO.OfficeColor.FromRgb(22, 163, 74), spacingBefore: global::OfficeIMO.OfficeLength.Points(8), spacingAfter: global::OfficeIMO.OfficeLength.Points(8))
                                .AddParagraph(p => {
                                    p.AddText("Ready: ").SetBold();
                                    p.AddText("row gutters are part of the composition model instead of caller-managed whitespace.");
                                }));
                            row.PercentColumn(50, column => column
                                .H2("Next")
                                .AddParagraph(p => p.AddText("The right column uses the same page flow but starts after an explicit gutter, giving report layouts a professional reading rhythm."))
                                .AddNumberedList(new[] {
                                    "Compose column content.",
                                    "Render each list item independently.",
                                    "Compare the raster baseline."
                                }, color: global::OfficeIMO.OfficeColor.FromRgb(55, 65, 81))
                                .AddTable(new[] {
                                    new[] { "Metric", "Value" },
                                    new[] { "Gutter", "18 pt" },
                                    new[] { "Panels", "Yes" }
                                }, style: new PdfTableStyle {
                                    HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85),
                                    HeaderTextColor = global::OfficeIMO.OfficeColor.White,
                                    RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                                    BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5),
                                    CellPaddingX = global::OfficeIMO.OfficeLength.Points(4),
                                    CellPaddingY = global::OfficeIMO.OfficeLength.Points(3),
                                    HeaderRowCount = 1,
                                    RightAlignNumeric = false,
                                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(7),
                                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(7),
                                    Alignments = new System.Collections.Generic.List<OfficeAlignment> { OfficeAlignment.Left, OfficeAlignment.Right },
                                    ColumnWidthWeights = new System.Collections.Generic.List<double> { 1.2, 0.8 }
                                })
                                .HR(global::OfficeIMO.OfficeLength.Points(0.8), global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207), global::OfficeIMO.OfficeLength.Points(8), global::OfficeIMO.OfficeLength.Points(8))
                                .AddParagraph(p => {
                                    p.AddText("Guarded: ").SetBold();
                                    p.AddText("the Poppler baseline catches cramped columns and accidental gutter regressions.");
                                }));
                        });
                        content.Column(column => {
                            column.Item().AddParagraph(p => p.AddText("End of row column sample."), OfficeAlignment.Right, global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80));
                        });
                    });
                });
            })
            .ToBytes();
    }

    private static byte[] CreateHeadersFooters() {
        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf header/footer gate", Alignment = OfficeAlignment.Left, Offset = global::OfficeIMO.OfficeLength.Points(18), Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Center, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Headers and Footers", author: "OfficeIMO")
            .H1("Header and Footer Baseline", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("Page one protects header placement, footer placement, and page number rendering on the first page."))
            .PanelParagraph(
                p => p.AddText("The same options should continue to render consistently after an explicit page break."),
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                    PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                    PaddingY = global::OfficeIMO.OfficeLength.Points(7)
                })
            .PageBreak()
            .H2("Continuation Page", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("Page two proves the raster harness compares more than the first page and guards the {page}/{pages} footer tokens."))
            .Paragraph(p => p.AddText("Right-aligned continuation note."), OfficeAlignment.Right, global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80))
            .ToBytes();
    }

    private static byte[] CreateStyleCheatsheet() {
        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf style cheatsheet", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Style Cheatsheet", author: "OfficeIMO")
            .H1("Style Cheatsheet", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .PanelParagraph(
                p => p.AddText("A compact visual sample for rich text, color, underline, and alignment behavior."),
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                    PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                    PaddingY = global::OfficeIMO.OfficeLength.Points(7)
                })
            .Paragraph(p => {
                p.AddText("Normal ");
                p.AddText("Bold ").SetBold();
                p.AddText("Italic ").SetItalic();
                p.AddText("Bold").SetBold();
                p.AddText(" Italic ").SetItalic();
                p.AddText("Underline ").SetUnderline();
            })
            .Paragraph(p => {
                p.AddText("Colors: ");
                p.AddText("Red ").SetColor(global::OfficeIMO.OfficeColor.FromRgb(200, 0, 0));
                p.AddText("Blue ").SetColor(global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180));
                p.AddText("Green").SetColor(global::OfficeIMO.OfficeColor.FromRgb(0, 128, 0));
            })
            .Paragraph(p => {
                p.AddText("Combinations: ");
                p.AddText("Bold ").SetBold();
                p.AddText("Italic ").SetItalic();
                p.AddText("Underlined ").SetUnderline();
                p.AddText("Bold ").SetBold();
                p.AddText("Italic ").SetItalic();
                p.AddText("Underlined ").SetUnderline();
            })
            .Paragraph(p => {
                p.AddText("Stateful toggles: ");
                p.AddText("bold on ").SetBold();
                p.AddText("bold off ");
                p.AddText("italic on ").SetItalic();
                p.AddText("italic off ");
                p.AddText("ul on ").SetUnderline();
                p.AddText("ul off");
            })
            .HR(0.8, global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207), 8, 8)
            .Paragraph(p => p.AddText("Center aligned line"), OfficeAlignment.Center)
            .Paragraph(p => p.AddText("Right aligned line"), OfficeAlignment.Right)
            .ToBytes();
    }

    private static byte[] CreateTableStyleGallery() {
        var rows = new[] {
            new[] { "Signal", "State", "Notes" },
            new[] { "Flow", "Generic", "Borders and row separators should reveal each preset shape at raster level." }
        };

        PdfDocument doc = CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9.5),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf Word-like table styles", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Table Style Gallery", author: "OfficeIMO")
            .H1("Table Style Gallery", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("Generic Word-style table names rendered by OfficeIMO.Pdf without invoice or report-specific behavior."));

        string[] visualStyleNames = {
            "TableGrid",
            "TableGridLight",
            "PlainTable1",
            "GridTable1Light",
            "ListTable1Light"
        };

        foreach (string styleName in visualStyleNames) {
            PdfTableStyle style = PdfTableStyles.FromWordTableStyle(styleName);
            style.Caption = styleName;
            style.CaptionColor = global::OfficeIMO.OfficeColor.FromRgb(80, 90, 100);
            style.CaptionFontSize = global::OfficeIMO.OfficeLength.Points(8.5);
            style.CaptionSpacingAfter = global::OfficeIMO.OfficeLength.Points(4);
            style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(6);
            style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(6);
            style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(76), global::OfficeIMO.OfficeLength.Points(70), global::OfficeIMO.OfficeLength.Points(300) };
            style.AutoFitColumns = false;
            style.Alignments = new List<OfficeAlignment> {
                OfficeAlignment.Left,
                OfficeAlignment.Center,
                OfficeAlignment.Left
            };

            doc.Table(rows, OfficeAlignment.Left, style);
        }

        doc.Table(CreateWordAccentSwatchRows(), OfficeAlignment.Left, CreateWordAccentSwatchStyle());

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
        var style = PdfTableStyles.ListTable1Light();
        style.Caption = "Accent swatches";
        style.CaptionColor = global::OfficeIMO.OfficeColor.FromRgb(80, 90, 100);
        style.CaptionFontSize = global::OfficeIMO.OfficeLength.Points(8.5);
        style.CaptionSpacingAfter = global::OfficeIMO.OfficeLength.Points(4);
        style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(6);
        style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(4);
        style.FontSize = global::OfficeIMO.OfficeLength.Points(8.5);
        style.HeaderFontSize = global::OfficeIMO.OfficeLength.Points(8.5);
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(5);
        style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(3);
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(76), global::OfficeIMO.OfficeLength.Points(45), global::OfficeIMO.OfficeLength.Points(45), global::OfficeIMO.OfficeLength.Points(45), global::OfficeIMO.OfficeLength.Points(45), global::OfficeIMO.OfficeLength.Points(45), global::OfficeIMO.OfficeLength.Points(45) };
        style.AutoFitColumns = false;
        style.Alignments = new List<OfficeAlignment> {
            OfficeAlignment.Left,
            OfficeAlignment.Center,
            OfficeAlignment.Center,
            OfficeAlignment.Center,
            OfficeAlignment.Center,
            OfficeAlignment.Center,
            OfficeAlignment.Center
        };

        var cellFills = new Dictionary<(int Row, int Column), global::OfficeIMO.OfficeColor>();
        for (int accent = 1; accent <= 6; accent++) {
            PdfTableStyle grid = PdfTableStyles.FromWordTableStyle("GridTable1LightAccent" + accent.ToString(System.Globalization.CultureInfo.InvariantCulture));
            PdfTableStyle list = PdfTableStyles.FromWordTableStyle("ListTable1LightAccent" + accent.ToString(System.Globalization.CultureInfo.InvariantCulture));
            cellFills[(1, accent)] = grid.BorderColor ?? global::OfficeIMO.OfficeColor.White;
            cellFills[(2, accent)] = grid.HeaderSeparatorColor ?? global::OfficeIMO.OfficeColor.White;
            cellFills[(3, accent)] = list.RowStripeFill ?? global::OfficeIMO.OfficeColor.White;
        }

        style.CellFills = cellFills;
        return style;
    }
}
