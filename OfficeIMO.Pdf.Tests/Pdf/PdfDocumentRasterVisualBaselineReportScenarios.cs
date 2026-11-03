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
    private static byte[] CreateLinksAndRules() {
        var rows = new[] {
            new[] { PdfTableCell.TextCell("Site"), PdfTableCell.TextCell("Label"), PdfTableCell.TextCell("Notes") },
            new[] { PdfTableCell.TextCell("OfficeIMO", "https://officeimo.net/"), PdfTableCell.TextCell("Homepage", "https://officeimo.net/"), PdfTableCell.TextCell("Docs", "https://officeimo.net/docs") },
            new[] { PdfTableCell.TextCell("GitHub", "https://github.com/EvotecIT/OfficeIMO"), PdfTableCell.TextCell("Repo", "https://github.com/EvotecIT/OfficeIMO"), PdfTableCell.TextCell("Issues", "https://github.com/EvotecIT/OfficeIMO/issues") }
        };

        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf links and rules", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Links and Rules", author: "OfficeIMO")
            .H1("Links & Rules Demo", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85), linkUri: "https://github.com/EvotecIT/OfficeIMO")
            .Paragraph(p => {
                p.AddText("Visit ");
                p.AddLink("OfficeIMO GitHub", "https://github.com/EvotecIT/OfficeIMO", global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180));
                p.AddText(" and the ");
                p.AddLink("project website", "https://officeimo.net/", global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180));
                p.AddText(" for more details.");
            })
            .HR(0.8, global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207), 8, 8)
            .Table(rows, OfficeAlignment.Left, CreateLinksTableStyle())
            .PanelParagraph(
                p => {
                    p.AddText("You can also place links ");
                    p.AddLink("inside panels", "https://officeimo.net/", global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180));
                    p.AddText(".");
                },
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                    BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5),
                    PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                    PaddingY = global::OfficeIMO.OfficeLength.Points(7)
                })
            .ToBytes();
    }

    private static byte[] CreateProfessionalReport() {
        string logoPath = Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png");
        byte[] logo = File.Exists(logoPath) ? File.ReadAllBytes(logoPath) : CreateFallbackLogo();
        byte[] alphaBadge = CreateTransparentBadgePng();

        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf professional report", Alignment = OfficeAlignment.Left, Offset = global::OfficeIMO.OfficeLength.Points(18), Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf professional report - page {page}/{pages}", Alignment = OfficeAlignment.Center, Show = true }
                },
                CreateOutlineFromHeadings = true
            })
            .Meta(title: "OfficeIMO.Pdf Professional Report", author: "OfficeIMO")
            .H1("Executive Security Summary", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => {
                p.AddText("A polished first-party PDF report built with ");
                p.AddText("OfficeIMO.Pdf").SetBold();
                p.AddText(", shared drawing descriptors, wrapped tables, images, and generated bookmarks.");
            })
            .Image(logo, 120, 48, OfficeAlignment.Right, fit: OfficeImageFit.Contain)
            .PanelParagraph(
                p => { p.AddText("Healthy").SetBold(); p.AddText(" posture with no critical findings."); },
                CreateStatusPanelStyle(),
                color: global::OfficeIMO.OfficeColor.FromRgb(42, 132, 82))
            .Image(alphaBadge, 18, 18, OfficeAlignment.Left)
            .Shape(CreateAccentRibbon(), spacingBefore: 12, spacingAfter: 10)
            .PanelParagraph(
                p => p.AddText("Operator note: long values should wrap cleanly, tables should stay inside the page, and reusable drawing primitives should remain available to Word, Excel, PowerPoint, and PDF exporters."),
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                    PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                    PaddingY = global::OfficeIMO.OfficeLength.Points(7)
                })
            .Table(new[] {
                new[] { "Signal", "Evidence", "Action" },
                new[] { "DMARC", "Policy is enforced for the primary domain and aligned subdomains.", "Monitor" },
                new[] { "TLS", "Certificate chain and protocol posture are ready for automated PSWriteOffice reports.", "Keep" },
                new[] { "DNS", "Delegation and stale-record checks are summarized without overflowing table cells.", "Review" },
                new[] { "PDF", "The report is generated by the MIT licensed dependency-free OfficeIMO.Pdf engine.", "Expand" }
            }, style: CreateReportTableStyle())
            .Image(logo, 112, 36, OfficeAlignment.Center, fit: OfficeImageFit.Contain)
            .Paragraph(p => p.AddText("Generated by OfficeIMO.Pdf."), OfficeAlignment.Right, global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80))
            .ToBytes();
    }

    internal static byte[] CreateLineItemsTwoPage() {
        string logoPath = Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png");
        byte[] logo = File.Exists(logoPath) ? File.ReadAllBytes(logoPath) : CreateFallbackLogo();
        var lineItemRows = CreateLineItemRows();
        var lineItemStyle = CreateLineItemGateTableStyle();
        var totalsRows = new[] {
            new[] { "Subtotal", "5 201,32 PLN" },
            new[] { "VAT 23%", "1 196,30 PLN" },
            new[] { "Total", "6 397,62 PLN" }
        };

        return CreateVisualBaselineDocument(new PdfCreateOptions {
                PageSetup = {
                    Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(595), global::OfficeIMO.OfficeLength.Points(842)),
                    Margins = { Left = global::OfficeIMO.OfficeLength.Points(50), Right = global::OfficeIMO.OfficeLength.Points(50), Top = global::OfficeIMO.OfficeLength.Points(54), Bottom = global::OfficeIMO.OfficeLength.Points(58) },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(11), Format = "{page} / {pages}", Alignment = OfficeAlignment.Center, Offset = global::OfficeIMO.OfficeLength.Points(28), Show = true }
                },
                BackgroundColor = null,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(11),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(25, 25, 25)
            })
            .Meta(title: "OfficeIMO.Pdf Line-Item Visual Gate", author: "OfficeIMO")
            .Compose(document => {
                document.Page(page => {
                    page.Content(content => {
                        content.Row(row => row
                            .PercentColumn(58, column => column
                                .H1("Statement #4048")
                                .AddParagraph(p => {
                                    p.AddText("Prepared: 23/12/2025\n");
                                    p.AddText("Review by: 06/01/2026");
                                },
                                    style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(2) }))
                            .PercentColumn(42, column => column
                                .AddImage(logo, global::OfficeIMO.OfficeLength.Points(156), global::OfficeIMO.OfficeLength.Points(54), OfficeAlignment.Right, fit: OfficeImageFit.Contain)));

                        content.Column(column => column
                            .Item()
                            .Spacer(global::OfficeIMO.OfficeLength.Points(41.4)));

                        content.Row(row => row
                            .Gap(global::OfficeIMO.OfficeLength.Points(30))
                            .PercentColumn(47, column => column
                                .H3("Prepared by")
                                .HR(global::OfficeIMO.OfficeLength.Points(1.2), global::OfficeIMO.OfficeColor.Black, spacingBefore: global::OfficeIMO.OfficeLength.Points(2), spacingAfter: global::OfficeIMO.OfficeLength.Points(10))
                                .AddParagraph(p => p.AddText("Syllabae Repraesentant\nOvum picem\nPrinceps avem distant, Linteum amicitia\nofficium21@aut statum.com\n881-306-3914"),
                                    style: new PdfParagraphStyle { LineHeight = 1.25, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }))
                            .PercentColumn(47, column => column
                                .H3("Recipient")
                                .HR(global::OfficeIMO.OfficeLength.Points(1.2), global::OfficeIMO.OfficeColor.Black, spacingBefore: global::OfficeIMO.OfficeLength.Points(2), spacingAfter: global::OfficeIMO.OfficeLength.Points(10))
                                .AddParagraph(p => p.AddText("Ceciderit Original\nAurum currunt\nSolis multum platea, Cocus fuge fluvio\nsubsisto93@celeritate.com\n839-621-9110"),
                                    style: new PdfParagraphStyle { LineHeight = 1.25, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })));

                        content.Column(column => {
                            column.Item()
                                .Spacer(global::OfficeIMO.OfficeLength.Points(33.4))
                                .AddTable(lineItemRows, OfficeAlignment.Left, lineItemStyle)
                                .AddTable(totalsRows, OfficeAlignment.Right, CreateLineItemTotalsTableStyle())
                                .AddPanelParagraph(
                                    p => { p.AddText("Document note: ").SetBold(); p.AddText("Project details, approval notes, or wrapper-provided metadata can be placed here by PSWriteOffice."); },
                                    new PdfPanelStyle {
                                        Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                                        BorderColor = global::OfficeIMO.OfficeColor.FromRgb(210, 218, 226),
                                        BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5),
                                        PaddingX = global::OfficeIMO.OfficeLength.Points(8),
                                        PaddingY = global::OfficeIMO.OfficeLength.Points(6),
                                        SpacingBefore = global::OfficeIMO.OfficeLength.Points(10),
                                        SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
                                    });
                        });
                    });
                });
            })
            .ToBytes();
    }

    private static PdfTableStyle CreateCoreTableStyle() {
        return new PdfTableStyle {
            HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(32, 76, 120),
            HeaderTextColor = global::OfficeIMO.OfficeColor.White,
            TextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
            RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(210, 218, 226),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5),
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(6),
            CellPaddingY = global::OfficeIMO.OfficeLength.Points(5),
            Caption = "Table 1. Core layout checks",
            CaptionColor = global::OfficeIMO.OfficeColor.FromRgb(80, 90, 100),
            CaptionFontSize = global::OfficeIMO.OfficeLength.Points(8.5),
            CaptionSpacingAfter = global::OfficeIMO.OfficeLength.Points(5),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(6),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(14),
            ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(82), global::OfficeIMO.OfficeLength.Points(310), global::OfficeIMO.OfficeLength.Points(76) },
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
        var style = PdfTableStyles.ListTable1Light();
        style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0);
        style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(14);
        style.Alignments = new List<OfficeAlignment> {
            OfficeAlignment.Right,
            OfficeAlignment.Left,
            OfficeAlignment.Right,
            OfficeAlignment.Right,
            OfficeAlignment.Right
        };
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(28), global::OfficeIMO.OfficeLength.Points(235), global::OfficeIMO.OfficeLength.Points(82), global::OfficeIMO.OfficeLength.Points(55), global::OfficeIMO.OfficeLength.Points(95) };
        style.AutoFitColumns = false;
        return style;
    }

    private static PdfTableStyle CreateLineItemTotalsTableStyle() {
        return new PdfTableStyle {
            HeaderFill = null,
            HeaderTextColor = global::OfficeIMO.OfficeColor.Black,
            TextColor = global::OfficeIMO.OfficeColor.FromRgb(25, 25, 25),
            FooterFill = null,
            FooterTextColor = global::OfficeIMO.OfficeColor.Black,
            RowStripeFill = null,
            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(224, 224, 224),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.4),
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(6),
            CellPaddingY = global::OfficeIMO.OfficeLength.Points(5),
            HeaderRowCount = 0,
            FooterRowCount = 1,
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(0),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(10),
            Alignments = new List<OfficeAlignment> {
                OfficeAlignment.Right,
                OfficeAlignment.Right
            },
            ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(105), global::OfficeIMO.OfficeLength.Points(105) }
        };
    }

    private static string FormatPln(decimal value) =>
        value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture).Replace('.', ',') + " PLN";

    private static PdfTableStyle CreateLineItemTableStyle() {
        return new PdfTableStyle {
            HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(32, 76, 120),
            HeaderTextColor = global::OfficeIMO.OfficeColor.White,
            TextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
            FooterFill = global::OfficeIMO.OfficeColor.FromRgb(232, 241, 248),
            FooterTextColor = global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85),
            RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(210, 218, 226),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5),
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(6),
            CellPaddingY = global::OfficeIMO.OfficeLength.Points(5),
            HeaderRowCount = 1,
            FooterRowCount = 1,
            Caption = "Table 1. Example line items",
            CaptionColor = global::OfficeIMO.OfficeColor.FromRgb(80, 90, 100),
            CaptionFontSize = global::OfficeIMO.OfficeLength.Points(8.5),
            CaptionSpacingAfter = global::OfficeIMO.OfficeLength.Points(5),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(6),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(14),
            RightAlignNumeric = true,
            Alignments = new List<OfficeAlignment> {
                OfficeAlignment.Left,
                OfficeAlignment.Right,
                OfficeAlignment.Right,
                OfficeAlignment.Right,
                OfficeAlignment.Left
            },
            ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(110), global::OfficeIMO.OfficeLength.Points(38), global::OfficeIMO.OfficeLength.Points(58), global::OfficeIMO.OfficeLength.Points(68), global::OfficeIMO.OfficeLength.Points(170) },
            AutoFitColumns = false
        };
    }

    private static PdfTableStyle CreateDashboardRiskTableStyle() {
        return new PdfTableStyle {
            HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42),
            HeaderTextColor = global::OfficeIMO.OfficeColor.White,
            RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.4),
            RowSeparatorColor = global::OfficeIMO.OfficeColor.FromRgb(226, 232, 240),
            RowSeparatorWidth = global::OfficeIMO.OfficeLength.Points(0.45),
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(7),
            CellPaddingY = global::OfficeIMO.OfficeLength.Points(5),
            HeaderFontSize = global::OfficeIMO.OfficeLength.Points(9.5),
            FontSize = global::OfficeIMO.OfficeLength.Points(9),
            RightAlignNumeric = true,
            ColumnWidthWeights = new List<double> { 2.1, 1.0, 0.8, 1.2 },
            Alignments = new List<OfficeAlignment> {
                OfficeAlignment.Left,
                OfficeAlignment.Center,
                OfficeAlignment.Right,
                OfficeAlignment.Left
            }
        };
    }

    private static PdfTableStyle CreateDashboardDecisionTableStyle() {
        return new PdfTableStyle {
            HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(30, 64, 175),
            HeaderTextColor = global::OfficeIMO.OfficeColor.White,
            RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(239, 246, 255),
            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(191, 219, 254),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.45),
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(7),
            CellPaddingY = global::OfficeIMO.OfficeLength.Points(5),
            HeaderRowCount = 1,
            FontSize = global::OfficeIMO.OfficeLength.Points(8.7),
            LineHeight = 1.15,
            ColumnWidthWeights = new List<double> { 1.0, 2.1 }
        };
    }

    private static PdfPanelStyle CreateStatusPanelStyle() {
        return new PdfPanelStyle {
            Background = global::OfficeIMO.OfficeColor.FromRgb(230, 247, 238),
            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(42, 132, 82),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(1.2),
            PaddingX = global::OfficeIMO.OfficeLength.Points(8),
            PaddingY = global::OfficeIMO.OfficeLength.Points(5),
            MaxWidth = global::OfficeIMO.OfficeLength.Points(245)
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
            HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(32, 76, 120),
            HeaderTextColor = global::OfficeIMO.OfficeColor.White,
            TextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
            RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(210, 218, 226),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5),
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(6),
            CellPaddingY = global::OfficeIMO.OfficeLength.Points(5),
            Caption = "Table 1. Report signals",
            CaptionColor = global::OfficeIMO.OfficeColor.FromRgb(80, 90, 100),
            CaptionFontSize = global::OfficeIMO.OfficeLength.Points(8.5),
            CaptionSpacingAfter = global::OfficeIMO.OfficeLength.Points(5),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(6),
            AutoFitColumns = true
        };
    }

    private static PdfTableStyle CreateLinksTableStyle() {
        return new PdfTableStyle {
            HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(32, 76, 120),
            HeaderTextColor = global::OfficeIMO.OfficeColor.White,
            TextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
            RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(210, 218, 226),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5),
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(6),
            CellPaddingY = global::OfficeIMO.OfficeLength.Points(5),
            Caption = "Table 1. Linked resources",
            CaptionColor = global::OfficeIMO.OfficeColor.FromRgb(80, 90, 100),
            CaptionFontSize = global::OfficeIMO.OfficeLength.Points(8.5),
            CaptionSpacingAfter = global::OfficeIMO.OfficeLength.Points(5),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(6),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(12)
        };
    }

    private static byte[] CreateFallbackLogo() => PdfPngTestImages.CreateRgbPng(255, 0, 0);

    private static byte[] CreateTransparentBadgePng() => PdfPngTestImages.CreateRgbaPng(0x2A, 0x84, 0x52, 0xA0);
}
