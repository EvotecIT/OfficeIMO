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

        return PdfDocument.Create(new PdfOptions {
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

        return PdfDocument.Create(new PdfOptions {
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

        return PdfDocument.Create(new PdfOptions {
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
