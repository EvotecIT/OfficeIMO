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
    private static byte[] CreateHelloWorld() {
        return CreateVisualBaselineDocument(new PdfOptions {
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
        return CreateVisualBaselineDocument(new PdfOptions {
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

        return CreateVisualBaselineDocument(new PdfOptions {
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

        return CreateVisualBaselineDocument(new PdfOptions {
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
        return CreateVisualBaselineDocument(new PdfOptions {
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

    private static byte[] CreateTabsLeaders() {
        return CreateVisualBaselineDocument(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf tabs and leaders",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true,
                DefaultParagraphStyle = new PdfParagraphStyle {
                    DefaultTabStopWidth = 252,
                    LineHeight = 1.18,
                    SpacingAfter = 4
                }
            })
            .Meta(title: "OfficeIMO.Pdf Tabs and Leaders", author: "OfficeIMO")
            .H1("Tabs and Leaders", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.Text("A compact visual gate for Word-like paragraph tabs, leader styles, and structured readback rhythm."))
            .HR(style: new PdfHorizontalRuleStyle {
                Color = PdfColor.FromRgb(183, 194, 207),
                Thickness = 0.8,
                SpacingBefore = 6,
                SpacingAfter = 8
            })
            .Paragraph(p => p.Bold("Revenue").Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right).Text("128 450"))
            .Paragraph(p => p.Text("Operating cost").Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right).Text("84 210"))
            .Paragraph(p => p.Text("Margin").Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right).Text("44 240"))
            .Paragraph(p => p.Text("Tax rate").Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.DecimalSeparator).Text("8.50"))
            .Paragraph(p => p.Text("Total").Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.DecimalSeparator).Text("1450.75"))
            .Paragraph(p => p.Text("Milestone").Tab(PdfTabLeaderStyle.Hyphens, PdfTabAlignment.Right).Text("Q4"))
            .Paragraph(p => p.Text("Signature").Tab(PdfTabLeaderStyle.Underscores, PdfTabAlignment.Left).Text("approved"))
            .Spacer(4)
            .Paragraph(p => p.Text("Left label").Tab().Text("plain tab").Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Center).Text("center"),
                style: new PdfParagraphStyle {
                    DefaultTabStopWidth = 144,
                    SpacingBefore = 4,
                    SpacingAfter = 2
                })
            .PanelParagraph(
                p => p.Text("Tab leaders are paragraph primitives, not invoice-specific rendering. Dotted value rows should align while remaining extractable as leader rows."),
                new PanelStyle {
                    Background = PdfColor.FromRgb(248, 250, 252),
                    BorderColor = PdfColor.FromRgb(183, 194, 207),
                    PaddingX = 9,
                    PaddingY = 7,
                    SpacingBefore = 10,
                    SpacingAfter = 8
                })
            .Paragraph(p => p.Text("End of tabs and leaders sample."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80))
            .ToBytes();
    }
}
