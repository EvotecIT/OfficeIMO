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
        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(11),
                PageSetup = {
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf hello-world - page {page}/{pages}", Alignment = OfficeAlignment.Center, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Hello World", author: "OfficeIMO")
            .H1("OfficeIMO.Pdf Hello World", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("This deterministic smoke page protects basic headings, paragraphs, metadata, and footer rendering."))
            .Paragraph(p => p.AddText("It is intentionally small so visual regressions in the simplest document path are easy to spot."))
            .ToBytes();
    }

    private static byte[] CreateCoreLayout() {
        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                PageSetup = {
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf core layout - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Core Layout", author: "OfficeIMO")
            .H1("Core Layout Baseline", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => {
                p.AddText("This page mixes ");
                p.AddText("rich text").SetBold();
                p.AddText(", colors, list flow, a panel, a horizontal rule, and a wrapped table.");
            },
                style: new PdfParagraphStyle { LineHeight = 1.5, LeftIndent = global::OfficeIMO.OfficeLength.Points(6), RightIndent = global::OfficeIMO.OfficeLength.Points(18), SpacingAfter = global::OfficeIMO.OfficeLength.Points(7) })
            .Bullets(new[] {
                "Natural proportional text spacing",
                "Stable vertical rhythm",
                "Readable table cells"
            }, color: global::OfficeIMO.OfficeColor.FromRgb(55, 65, 81))
            .Numbered(new[] {
                "Measure the generated document.",
                "Compare the approved raster image.",
                "Improve the engine instead of hiding rough output."
            }, color: global::OfficeIMO.OfficeColor.FromRgb(55, 65, 81))
            .PanelParagraph(
                p => { p.AddText("Panel check").SetBold(); p.AddLineBreak(); p.AddText("Text should sit comfortably inside its border and retain readable padding after rasterization."); },
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                    PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                    PaddingY = global::OfficeIMO.OfficeLength.Points(7)
                },
                color: global::OfficeIMO.OfficeColor.FromRgb(35, 88, 65))
            .HR(thickness: 1, color: global::OfficeIMO.OfficeColor.FromRgb(32, 76, 120), spacingBefore: 8, spacingAfter: 8)
            .Table(new[] {
                new[] { "Area", "Expectation", "State" },
                new[] { "Paragraphs", "Rich text wraps and keeps comfortable line spacing.", "Guarded" },
                new[] { "Lists", "Bullets and numbered steps align with wrapped text.", "Guarded" },
                new[] { "Tables", "Long text stays inside cells and body text remains readable after colored headers.", "Guarded" }
            }, style: CreateCoreTableStyle())
            .Paragraph(p => p.AddText("End of core layout baseline."), OfficeAlignment.Right, global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80))
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

        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf lists and tables", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Lists and Tables", author: "OfficeIMO")
            .H1("Lists and Tables", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("A compact report sample for list rhythm, numeric alignment, footer rows, and wrapped table notes."))
            .Bullets(new[] {
                "Report-friendly list spacing",
                "Right-aligned quantities and amounts",
                "Footer row that remains visually distinct"
            }, OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(55, 65, 81))
            .Numbered(new[] {
                "Collect line items.",
                "Render the table with stable column widths.",
                "Review the generated PDF through the raster gate."
            }, OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(55, 65, 81))
            .Table(rows, OfficeAlignment.Left, CreateLineItemTableStyle())
            .Paragraph(p => p.AddText("End of lists and tables sample."), OfficeAlignment.Right, global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80))
            .ToBytes();
    }

    private static byte[] CreateDefaultStyles() {
        var rows = new[] {
            new[] { "Metric", "Current", "Target" },
            new[] { "Runtime dependencies", "0", "0" },
            new[] { "Visual gates", "Growing", "Required for public claims" },
            new[] { "PowerShell wrapper", "PSWriteOffice", "Expose safe PDF operations" }
        };

        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf default styles", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                },
                DefaultTableStyle = PdfTableStyles.Light()
            })
            .Meta(title: "OfficeIMO.Pdf Default Styles", author: "OfficeIMO")
            .H1("Default Styles", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("This sample uses document-level defaults for text color, headers, footers, and the light table preset."))
            .PanelParagraph(
                p => p.AddText("The default table style should be good enough for a simple business report without every caller hand-tuning colors and padding."),
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                    PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                    PaddingY = global::OfficeIMO.OfficeLength.Points(7)
                })
            .Table(rows)
            .ToBytes();
    }

    private static byte[] CreateStyledRuns() {
        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf styled runs", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                }
            })
            .Meta(title: "OfficeIMO.Pdf Styled Runs", author: "OfficeIMO")
            .H1("Styled Runs", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("A compact visual sample for inline font style, color, underline, strike-through, and stateful run toggles."))
            .PanelParagraph(
                p => {
                    p.AddText("Inline styles").SetBold();
                    p.AddText(" should remain readable in normal business-report text, not only in synthetic text extraction checks.");
                },
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                    PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                    PaddingY = global::OfficeIMO.OfficeLength.Points(7)
                })
            .Paragraph(p => {
                p.AddText("You can mix ");
                p.AddText("bold ").SetBold();
                p.AddText("italic ").SetItalic();
                p.AddText("bold italic ").SetBold().SetItalic();
                p.AddText("underlined ").SetUnderline();
                p.AddText("obsolete ").SetStrikethrough();
                p.AddText("and ").SetColor(global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80));
                p.AddText("colors.").SetColor(global::OfficeIMO.OfficeColor.FromRgb(8, 28, 120));
            },
                style: new PdfParagraphStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(14) })
            .Paragraph(p => {
                p.AddText("Underline respects color: ");
                p.AddText("red").SetUnderline().SetColor(global::OfficeIMO.OfficeColor.FromRgb(200, 0, 0));
                p.AddText(", ");
                p.AddText("blue").SetUnderline().SetColor(global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180));
                p.AddText(", and ");
                p.AddText("green").SetUnderline().SetColor(global::OfficeIMO.OfficeColor.FromRgb(0, 128, 0));
                p.AddText(".");
            })
            .Paragraph(p => {
                p.AddText("Stateful color toggles: ");
                p.AddText("critical ").SetColor(global::OfficeIMO.OfficeColor.FromRgb(185, 28, 28));
                p.AddText("informational ").SetColor(global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180));
                p.AddText("healthy ").SetColor(global::OfficeIMO.OfficeColor.FromRgb(22, 101, 52));
                p.AddText("normal.").SetColor(global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55));
            })
            .Paragraph(p => p.AddText("End of styled runs sample."), OfficeAlignment.Right, global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80))
            .ToBytes();
    }

    private static byte[] CreateTabsLeaders() {
        return CreateVisualBaselineDocument(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                PageSetup = {
                    Header = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf tabs and leaders", Alignment = OfficeAlignment.Left, Show = true },
                    Footer = { StandardFont = PdfStandardFont.Helvetica, FontSize = global::OfficeIMO.OfficeLength.Points(8), Format = "OfficeIMO.Pdf examples - page {page}/{pages}", Alignment = OfficeAlignment.Right, Show = true }
                },
                DefaultParagraphStyle = new PdfParagraphStyle {
                    DefaultTabStopWidth = global::OfficeIMO.OfficeLength.Points(252),
                    LineHeight = 1.18,
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(4)
                }
            })
            .Meta(title: "OfficeIMO.Pdf Tabs and Leaders", author: "OfficeIMO")
            .H1("Tabs and Leaders", OfficeAlignment.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
            .Paragraph(p => p.AddText("A compact visual gate for Word-like paragraph tabs, leader styles, and structured readback rhythm."))
            .HR(style: new PdfHorizontalRuleStyle {
                Color = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                Thickness = global::OfficeIMO.OfficeLength.Points(0.8),
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(6),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(8)
            })
            .Paragraph(p => { p.AddText("Revenue").SetBold(); p.AddTab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right); p.AddText("128 450"); })
            .Paragraph(p => { p.AddText("Operating cost"); p.AddTab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right); p.AddText("84 210"); })
            .Paragraph(p => { p.AddText("Margin"); p.AddTab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right); p.AddText("44 240"); })
            .Paragraph(p => { p.AddText("Tax rate"); p.AddTab(PdfTabLeaderStyle.Dots, PdfTabAlignment.DecimalSeparator); p.AddText("8.50"); })
            .Paragraph(p => { p.AddText("Total"); p.AddTab(PdfTabLeaderStyle.Dots, PdfTabAlignment.DecimalSeparator); p.AddText("1450.75"); })
            .Paragraph(p => { p.AddText("Milestone"); p.AddTab(PdfTabLeaderStyle.Hyphens, PdfTabAlignment.Right); p.AddText("Q4"); })
            .Paragraph(p => { p.AddText("Signature"); p.AddTab(PdfTabLeaderStyle.Underscores, PdfTabAlignment.Left); p.AddText("approved"); })
            .Spacer(4)
            .Paragraph(p => { p.AddText("Left label"); p.AddTab(); p.AddText("plain tab"); p.AddTab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Center); p.AddText("center"); },
                style: new PdfParagraphStyle {
                    DefaultTabStopWidth = global::OfficeIMO.OfficeLength.Points(144),
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(4),
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(2)
                })
            .PanelParagraph(
                p => p.AddText("Tab leaders are paragraph primitives, not invoice-specific rendering. Dotted value rows should align while remaining extractable as leader rows."),
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                    PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                    PaddingY = global::OfficeIMO.OfficeLength.Points(7),
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(10),
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(8)
                })
            .Paragraph(p => p.AddText("End of tabs and leaders sample."), OfficeAlignment.Right, global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80))
            .ToBytes();
    }
}
