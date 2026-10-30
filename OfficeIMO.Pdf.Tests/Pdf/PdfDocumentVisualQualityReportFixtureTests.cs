using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void RowColumns_KeepTextInsideColumnFramesWithReadableRhythm() {
        const double pageWidth = 420;
        const double margin = 30;
        const double gutter = 24;
        double contentWidth = pageWidth - margin - margin;
        double columnWidth = (contentWidth - gutter) / 2;
        double leftX = margin;
        double leftRightX = leftX + columnWidth;
        double rightX = leftRightX + gutter;
        double rightRightX = rightX + columnWidth;

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(pageWidth),
                PageHeight = global::OfficeIMO.OfficeLength.Points(280),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(margin),
                MarginRight = global::OfficeIMO.OfficeLength.Points(margin),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Gap(global::OfficeIMO.OfficeLength.Points(gutter))
                                .PercentColumn(50, column => column
                                    .H2("LeftFlow")
                                    .Paragraph(p => p.Text("LeftAlphaOne carries enough ordinary report text to wrap inside its column without touching the neighboring frame."))
                                    .Bullets(new[] {
                                        "LeftBulletOne stays inside the measure.",
                                        "LeftBulletTwo keeps a clear baseline."
                                    })
                                    .PanelParagraph(
                                        p => p.Bold("LeftPanel").Text(": spacing remains visible after the list."),
                                        new PdfPanelStyle {
                                            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(191, 191, 191),
                                            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5),
                                            PaddingX = global::OfficeIMO.OfficeLength.Points(6),
                                            PaddingY = global::OfficeIMO.OfficeLength.Points(5),
                                            Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252)
                                        }))
                                .PercentColumn(50, column => column
                                    .H2("RightFlow")
                                    .Paragraph(p => p.Text("RightAlphaOne uses the same generic layout primitives and should start after the explicit gutter."))
                                    .Numbered(new[] {
                                        "RightStepOne composes content.",
                                        "RightStepTwo preserves reading rhythm."
                                    })
                                    .PanelParagraph(
                                        p => p.Bold("RightPanel").Text(": the final note avoids cramped text."),
                                        new PdfPanelStyle {
                                            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(191, 191, 191),
                                            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5),
                                            PaddingX = global::OfficeIMO.OfficeLength.Points(6),
                                            PaddingY = global::OfficeIMO.OfficeLength.Points(5),
                                            Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252)
                                        }))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var renderedPage = pdf.GetPage(1);
        var leftLines = GetVisualTextLines(renderedPage, leftX - 1, leftRightX + 1);
        var rightLines = GetVisualTextLines(renderedPage, rightX - 1, rightRightX + 1);

        Assert.Contains(leftLines, line => line.Text.Contains("LeftAlphaOne", StringComparison.Ordinal));
        Assert.Contains(rightLines, line => line.Text.Contains("RightAlphaOne", StringComparison.Ordinal));
        Assert.True(leftLines.Count >= 7, $"Expected the left flow to produce multiple visual lines. Lines: {leftLines.Count}.");
        Assert.True(rightLines.Count >= 7, $"Expected the right flow to produce multiple visual lines. Lines: {rightLines.Count}.");

        Assert.All(leftLines, line =>
            Assert.True(line.X1 >= leftX - 1 && line.X2 <= leftRightX + 1.5,
                $"Expected left column line '{line.Text}' to stay inside {leftX:0.##}..{leftRightX:0.##}, but it rendered at {line.X1:0.##}..{line.X2:0.##}."));
        Assert.All(rightLines, line =>
            Assert.True(line.X1 >= rightX - 1.5 && line.X2 <= rightRightX + 1.5,
                $"Expected right column line '{line.Text}' to stay inside {rightX:0.##}..{rightRightX:0.##}, but it rendered at {line.X1:0.##}..{line.X2:0.##}."));

        foreach (var leftLine in leftLines) {
            foreach (var rightLine in rightLines.Where(line => Math.Abs(line.BaselineY - leftLine.BaselineY) <= 0.2)) {
                double clearance = rightLine.X1 - leftLine.X2;
                Assert.True(clearance >= gutter - 1,
                    $"Expected row columns to preserve the {gutter:0.##}pt gutter between '{leftLine.Text}' and '{rightLine.Text}'. Clearance: {clearance:0.##}pt.");
            }
        }

        AssertReadableTextRhythm(leftLines, "left column");
        AssertReadableTextRhythm(rightLines, "right column");
    }

    [Fact]
    public void RowColumns_UseBuiltInWordLikeGutterWhenNoGapIsConfigured() {
        const double pageWidth = 360;
        const double margin = 30;
        double gutter = PdfRowStyle.DefaultGap.ToPoints();
        double contentWidth = pageWidth - margin - margin;
        double columnWidth = (contentWidth - gutter) / 2;
        double leftX = margin;
        double leftRightX = leftX + columnWidth;
        double rightX = leftRightX + gutter;
        double rightRightX = rightX + columnWidth;

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(pageWidth),
                PageHeight = global::OfficeIMO.OfficeLength.Points(180),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(margin),
                MarginRight = global::OfficeIMO.OfficeLength.Points(margin),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row => row
                            .PercentColumn(50, column => column.Paragraph(p => p.Text("LeftPlainColumn wraps in the first default column frame.")))
                            .PercentColumn(50, column => column.Paragraph(p => p.Text("RightPlainColumn starts after the built-in row gutter.")))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var renderedPage = pdf.GetPage(1);
        var leftLines = GetVisualTextLines(renderedPage, leftX - 1, leftRightX + 1);
        var rightLines = GetVisualTextLines(renderedPage, rightX - 1, rightRightX + 1);
        double rightStart = FindWordStartX(renderedPage, "RightPlainColumn");

        Assert.Contains(leftLines, line => line.Text.Contains("LeftPlainColumn", StringComparison.Ordinal));
        Assert.Contains(rightLines, line => line.Text.Contains("RightPlainColumn", StringComparison.Ordinal));
        Assert.True(rightStart >= rightX - 1,
            $"Expected an unstyled two-column row to use the built-in {gutter:0.##}pt gutter. Right column started at {rightStart:0.##}, expected at least {rightX:0.##}.");
        Assert.All(leftLines, line =>
            Assert.True(line.X1 >= leftX - 1 && line.X2 <= leftRightX + 1.5,
                $"Expected default-gutter left column line '{line.Text}' to stay inside the left column frame."));
        Assert.All(rightLines, line =>
            Assert.True(line.X1 >= rightX - 1.5 && line.X2 <= rightRightX + 1.5,
                $"Expected default-gutter right column line '{line.Text}' to stay inside the right column frame."));
        AssertNoSameBaselineTextCollisions(renderedPage, "default-gutter row columns");
    }

    [Fact]
    public void MixedWordLikeFlow_KeepsReadableRhythmAcrossGenericPrimitives() {
        const double pageWidth = 420;
        const double margin = 36;
        const double gutter = 24;
        double contentWidth = pageWidth - margin - margin;
        double columnWidth = (contentWidth - gutter) / 2;
        double leftX = margin;
        double leftRightX = leftX + columnWidth;
        double rightX = leftRightX + gutter;
        double rightRightX = rightX + columnWidth;
        byte[] png = CreateMinimalRgbPng();
        var shape = OfficeShape.Rectangle(72, 16);
        shape.FillColor = OfficeColor.LightBlue;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 0.75;

        var paragraphStyle = new PdfParagraphStyle {
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(10),
            LineHeight = 1.25
        };
        var panelStyle = new PdfPanelStyle {
            Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(191, 191, 191),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5),
            PaddingX = global::OfficeIMO.OfficeLength.Points(8),
            PaddingY = global::OfficeIMO.OfficeLength.Points(6),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(2),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(10)
        };
        var listStyle = new PdfListStyle {
            LeftIndent = global::OfficeIMO.OfficeLength.Points(10),
            MarkerGap = global::OfficeIMO.OfficeLength.Points(6),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(10),
            ItemSpacing = global::OfficeIMO.OfficeLength.Points(2)
        };
        var tableStyle = PdfTableStyles.Minimal();
        tableStyle.HeaderRowCount = 0;
        tableStyle.CellPaddingX = global::OfficeIMO.OfficeLength.Points(6);
        tableStyle.CellPaddingY = global::OfficeIMO.OfficeLength.Points(4);
        tableStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(2);
        tableStyle.SpacingAfter = global::OfficeIMO.OfficeLength.Points(10);
        var imageStyle = new PdfImageStyle {
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(4),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(10)
        };
        var drawingStyle = new PdfDrawingStyle {
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(2),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(10)
        };

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(pageWidth),
                PageHeight = global::OfficeIMO.OfficeLength.Points(620),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(margin),
                MarginRight = global::OfficeIMO.OfficeLength.Points(margin),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content => {
                        content.Column(column => column.Item()
                            .H1("WordFlowGate", new PdfHeadingStyle {
                                FontSize = global::OfficeIMO.OfficeLength.Points(18),
                                SpacingAfter = global::OfficeIMO.OfficeLength.Points(12)
                            })
                            .Paragraph(p => p.Text("LeadMarker introduces a generic report section without any invoice-specific shape."), style: paragraphStyle)
                            .PanelParagraph(p => p.Bold("PanelMarker").Text(": boxed notes keep visible padding and downstream rhythm."), panelStyle)
                            .Bullets(new[] {
                                "BulletMarker keeps list text in the normal document flow.",
                                "BulletSecond keeps a second readable baseline."
                            }, style: listStyle)
                            .Table(new[] {
                                new[] { "TableMarker", "Ready" },
                                new[] { "Rhythm", "Stable" }
                            }, style: tableStyle)
                            .Image(png, global::OfficeIMO.OfficeLength.Points(24), global::OfficeIMO.OfficeLength.Points(24), style: imageStyle)
                            .Shape(shape, style: drawingStyle)
                            .Paragraph(p => p.Text("AfterVisualMarker follows image and shape blocks with deliberate breathing room."), style: paragraphStyle));

                        content.Row(row => row
                            .Gap(global::OfficeIMO.OfficeLength.Points(gutter))
                            .Style(new PdfRowStyle {
                                SpacingBefore = global::OfficeIMO.OfficeLength.Points(4),
                                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
                            })
                            .PercentColumn(50, column => column
                                .H2("LeftMixed")
                                .Paragraph(p => p.Text("LeftMixedMarker wraps safely inside the left column measure."), style: paragraphStyle)
                                .Bullets(new[] { "LeftMixedBullet keeps rhythm." }, style: listStyle))
                            .PercentColumn(50, column => column
                                .H2("RightMixed")
                                .Paragraph(p => p.Text("RightMixedMarker starts after the explicit row gutter."), style: paragraphStyle)
                                .Numbered(new[] { "RightMixedStep keeps rhythm." }, style: listStyle)));
                    })))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        string pageContent = string.Join("\n", GetPageContentStreams(bytes, 1));
        double leadY = FindWordStartY(page, "LeadMarker");
        double panelY = FindWordStartY(page, "PanelMarker");
        double bulletY = FindWordStartY(page, "BulletMarker");
        double tableY = FindWordStartY(page, "TableMarker");
        double afterVisualY = FindWordStartY(page, "AfterVisualMarker");
        var leftLines = GetVisualTextLines(page, leftX - 1, leftRightX + 1);
        var rightLines = GetVisualTextLines(page, rightX - 1, rightRightX + 1);
        var pageLines = GetVisualTextLines(page, 0, pageWidth);

        Assert.Equal(1, pdf.NumberOfPages);
        Assert.True(leadY - panelY >= 18, $"Expected panel content to sit below the lead paragraph with readable rhythm. Gap: {leadY - panelY:0.##}pt.");
        Assert.True(panelY - bulletY >= 18, $"Expected list content to sit below the panel with readable rhythm. Gap: {panelY - bulletY:0.##}pt.");
        Assert.True(bulletY - tableY >= 18, $"Expected table content to sit below the list with readable rhythm. Gap: {bulletY - tableY:0.##}pt.");
        Assert.True(tableY - afterVisualY >= 55, $"Expected text after image and shape blocks to preserve visual breathing room. Gap: {tableY - afterVisualY:0.##}pt.");
        Assert.Contains("/Im1 Do", pageContent);
        Assert.Contains("72 16 re B", pageContent);
        Assert.Contains(leftLines, line => line.Text.Contains("LeftMixedMarker", StringComparison.Ordinal));
        Assert.Contains(rightLines, line => line.Text.Contains("RightMixedMarker", StringComparison.Ordinal));
        Assert.All(leftLines, line =>
            Assert.True(line.X1 >= leftX - 1 && line.X2 <= leftRightX + 1.5,
                $"Expected mixed left-column line '{line.Text}' to stay inside the left column frame."));
        Assert.All(rightLines, line =>
            Assert.True(line.X1 >= rightX - 1.5 && line.X2 <= rightRightX + 1.5,
                $"Expected mixed right-column line '{line.Text}' to stay inside the right column frame."));
        AssertReadableTextRhythm(leftLines.Where(line => line.Text.Contains("Mixed", StringComparison.Ordinal)).ToList(), "mixed left column");
        AssertReadableTextRhythm(rightLines.Where(line => line.Text.Contains("Mixed", StringComparison.Ordinal)).ToList(), "mixed right column");
        AssertNoCrampedBaselines(pageLines, "mixed Word-like flow");
        AssertNoSameBaselineTextCollisions(page, "mixed Word-like flow");
        AssertNoAmbiguousSameBaselineRunGaps(page, "mixed Word-like flow");
    }

    [Fact]
    public void WordLikeLineItemTable_KeepsReadableColumnsWithoutTemplateApi() {
        const double pageWidth = 595;
        const double margin = 56;
        var style = PdfTableStyles.ListTable1Light();
        style.RightAlignNumeric = true;
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(5);
        style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(5.5);
        style.ColumnWidthWeights = new List<double> { 0.45, 4.2, 1.15, 0.8, 1.2 };
        style.ColumnMinWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(22), global::OfficeIMO.OfficeLength.Points(185), global::OfficeIMO.OfficeLength.Points(62), global::OfficeIMO.OfficeLength.Points(34), global::OfficeIMO.OfficeLength.Points(68) };
        style.FooterRowCount = 1;
        style.FooterSeparatorColor = global::OfficeIMO.OfficeColor.Black;
        style.FooterSeparatorWidth = global::OfficeIMO.OfficeLength.Points(0.9);
        style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(8);
        style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(10);

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(pageWidth),
                PageHeight = global::OfficeIMO.OfficeLength.Points(760),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(margin),
                MarginRight = global::OfficeIMO.OfficeLength.Points(margin),
                MarginTop = global::OfficeIMO.OfficeLength.Points(56),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(56)
            })
            .Theme(PdfTheme.WordLike())
            .H1("LineItemGate")
            .Paragraph(p => p.Text("A generic Word-like line item table protects table rhythm without adding invoice-specific engine APIs."))
            .Table(new[] {
                new[] { "#", "Product", "UnitPrice", "Qty", "Total" },
                new[] { "1", "MonitoringSeats", "31.80", "2", "63.60" },
                new[] { "2", "RadioInsulamPluviae", "62.57", "7", "437.99" },
                new[] { "3", "Long Wrapping Service Description For Column Rhythm", "42.50", "5", "212.50" },
                new[] { "4", "RexMaximeDixitque", "22.75", "5", "113.75" },
                new[] { "5", "ActumExemplumPrinceps", "6.41", "8", "51.28" },
                new[] { "6", "CustodiPuella", "79.05", "8", "632.40" },
                new[] { "", "TotalDue", "", "", "1499.12" }
            }, style: style)
            .Paragraph(p => p.Text("LineItemGateEnd"), PdfAlign.Center)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var allLines = GetVisualTextLines(page, 0, pageWidth);
        double tableStartY = FindWordStartY(page, "Product");
        double tableEndY = FindWordStartY(page, "TotalDue");
        var tableLines = allLines
            .Where(line => line.BaselineY <= tableStartY + 1 && line.BaselineY >= tableEndY - 1)
            .ToList();

        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains(tableLines, line => line.Text.Contains("MonitoringSeats", StringComparison.Ordinal));
        Assert.Contains(tableLines, line => line.Text.Contains("LongWrappingService", StringComparison.Ordinal));
        Assert.Contains(tableLines, line => line.Text.Contains("Description", StringComparison.Ordinal));
        Assert.Contains(tableLines, line => line.Text.Contains("TotalDue", StringComparison.Ordinal));

        Assert.True(FindWordEndX(page, "MonitoringSeats") < FindWordStartX(page, "31.80") - 10,
            "Expected product text to end with visible space before the unit price column.");
        double monitoringBaselineY = FindWordStartY(page, "MonitoringSeats");
        Assert.True(FindWordEndXOnBaseline(page, "31.80", monitoringBaselineY) < FindWordStartXOnBaseline(page, "2", monitoringBaselineY) - 10,
            "Expected unit price text to stay separated from the quantity column.");
        Assert.True(FindWordEndXOnBaseline(page, "2", monitoringBaselineY) < FindWordStartXOnBaseline(page, "63.60", monitoringBaselineY) - 10,
            "Expected quantity text to stay separated from the total column.");
        double footerBaselineY = FindWordStartY(page, "TotalDue");
        Assert.True(FindWordEndXOnBaseline(page, "TotalDue", footerBaselineY) < FindWordStartXOnBaseline(page, "1499.12", footerBaselineY) - 10,
            "Expected footer label and numeric summary to stay visibly separated.");
        Assert.True(FindWordEndX(page, "632.40") <= pageWidth - margin + 1,
            "Expected the rightmost total to stay inside the document margin.");
        Assert.True(FindWordEndX(page, "1499.12") <= pageWidth - margin + 1,
            "Expected the footer total to stay inside the document margin.");
        Assert.True(FindWordStartY(page, "LineItemGateEnd") < tableEndY - 20,
            "Expected following content to retain breathing room after the table.");

        AssertNoCrampedBaselines(tableLines, "generic line item table");
        AssertNoSameBaselineTextCollisions(page, "generic line item table");
        AssertNoAmbiguousSameBaselineRunGaps(page, "generic line item table");
    }

    [Fact]
    public void TwoPageLineItemStatementFixture_KeepsReadableReportRhythmWithoutTemplateApi() {
        const double pageWidth = 595;
        const double marginLeft = 50;
        const double marginRight = 50;
        double contentRight = pageWidth - marginRight;

        byte[] bytes = PdfDocumentRasterVisualBaselineTests.CreateLineItemsTwoPage();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);

        var page1 = pdf.GetPage(1);
        var page1Lines = GetVisualTextLines(page1, 0, pageWidth);
        double titleY = FindWordStartY(page1, "Statement");
        double preparedDateY = FindWordStartY(page1, "Prepared:");
        double preparedHeadingY = FindWordStartY(page1, "Preparedby");
        double tableHeaderY = FindWordStartY(page1, "Product");
        double firstRowY = FindWordStartY(page1, "Experientiam");
        double secondRowY = FindWordStartY(page1, "Radio");
        double lastPage1RowY = FindWordStartY(page1, "Custodi");

        Assert.True(titleY - preparedDateY >= 22,
            $"Expected issue metadata to sit comfortably below the statement title. Gap: {titleY - preparedDateY:0.##}pt.");
        Assert.True(preparedDateY - preparedHeadingY >= 56,
            $"Expected sender/recipient blocks to start after the header with visible breathing room. Gap: {preparedDateY - preparedHeadingY:0.##}pt.");
        Assert.True(preparedHeadingY - tableHeaderY >= 118,
            $"Expected the line-item table to start after the address blocks, not collide with them. Gap: {preparedHeadingY - tableHeaderY:0.##}pt.");
        Assert.True(tableHeaderY - firstRowY >= 15,
            $"Expected table header and first row to retain readable rhythm. Gap: {tableHeaderY - firstRowY:0.##}pt.");
        Assert.True(firstRowY - secondRowY >= 17,
            $"Expected body rows to keep readable baseline rhythm. Gap: {firstRowY - secondRowY:0.##}pt.");

        AssertStatementRowColumns(page1, "Experientiam", "31,80", "2", "63,60", "page 1 first row");
        AssertStatementRowColumns(page1, "Custodi", "79,05", "8", "632,40", "page 1 last visible row");
        Assert.True(FindWordStartX(page1, "Product") >= marginLeft + 24,
            "Expected the product header to sit inside the line-item table frame.");
        Assert.True(FindWordEndX(page1, "632,40") <= contentRight + 1,
            "Expected the rightmost page 1 total to stay inside the document margin.");
        Assert.True(lastPage1RowY > 78,
            $"Expected the final page 1 row to leave room for the footer. Baseline: {lastPage1RowY:0.##}pt.");

        AssertNoCrampedBaselines(page1Lines, "two-page statement page 1");
        AssertNoSameBaselineTextCollisions(page1, "two-page statement page 1");
        AssertNoAmbiguousSameBaselineRunGaps(page1, "two-page statement page 1");

        var page2 = pdf.GetPage(2);
        var page2Lines = GetVisualTextLines(page2, 0, pageWidth);
        double page2FirstRowY = FindWordStartY(page2, "Praestare");
        double page2SecondRowY = FindWordStartY(page2, "Umero");
        double page2LastItemY = FindWordStartY(page2, "Finis");
        double subtotalY = FindWordStartY(page2, "Subtotal");
        double vatY = FindWordStartY(page2, "VAT");
        double totalValueY = FindWordStartY(page2, "6397,62");
        double noteY = FindWordStartY(page2, "Documentnote:");

        Assert.True(page2FirstRowY - page2SecondRowY >= 17,
            $"Expected continued body rows to keep readable baseline rhythm. Gap: {page2FirstRowY - page2SecondRowY:0.##}pt.");
        Assert.True(page2LastItemY - subtotalY >= 24,
            $"Expected summary totals to sit after the final line item with visible breathing room. Gap: {page2LastItemY - subtotalY:0.##}pt.");
        Assert.True(subtotalY - vatY >= 14,
            $"Expected totals rows to keep readable rhythm. Gap: {subtotalY - vatY:0.##}pt.");
        Assert.True(vatY - totalValueY >= 14,
            $"Expected VAT and total rows to stay separated. Gap: {vatY - totalValueY:0.##}pt.");
        Assert.True(totalValueY - noteY >= 22,
            $"Expected document note panel to follow totals with breathing room. Gap: {totalValueY - noteY:0.##}pt.");

        AssertStatementRowColumns(page2, "Umero", "81,72", "2", "163,44", "page 2 continued row");
        double subtotalBaselineY = FindWordStartY(page2, "Subtotal");
        Assert.True(FindWordEndXOnBaseline(page2, "Subtotal", subtotalBaselineY) < FindWordStartXOnBaseline(page2, "5201,32", subtotalBaselineY) - 10,
            "Expected subtotal label and value to stay visibly separated.");
        Assert.True(FindWordEndX(page2, "6397,62") <= contentRight + 1,
            "Expected the rightmost grand total to stay inside the document margin.");

        AssertNoCrampedBaselines(page2Lines, "two-page statement page 2");
        AssertNoSameBaselineTextCollisions(page2, "two-page statement page 2");
        AssertNoAmbiguousSameBaselineRunGaps(page2, "two-page statement page 2");
    }

    [Fact]
    public void ShowcaseDashboard_KeepsReadableGenericLayoutGeometry() {
        const double pageWidth = 841.89;
        const double marginLeft = 42;
        const double marginRight = 42;
        const double bodyGap = 18;
        double contentWidth = pageWidth - marginLeft - marginRight;
        double bodyColumnWidth = contentWidth - bodyGap;
        double leftColumnRightX = marginLeft + (bodyColumnWidth * 0.58);
        double rightColumnX = leftColumnRightX + bodyGap;
        double rightColumnRightX = pageWidth - marginRight;

        byte[] bytes = PdfDocumentRasterVisualBaselineTests.CreateShowcaseDashboard();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var leftLines = GetVisualTextLines(page, marginLeft - 1, leftColumnRightX + 1);
        var rightLines = GetVisualTextLines(page, rightColumnX - 1, rightColumnRightX + 1);

        double titleY = FindWordStartY(page, "Quarterly");
        double leadY = FindWordStartY(page, "single-page");
        double firstMetricY = FindWordStartY(page, "92%");
        double secondMetricY = FindWordStartY(page, "1.8h");
        double thirdMetricY = FindWordStartY(page, "34");
        double fourthMetricY = FindWordStartY(page, "Critical");
        double leftHeadingY = FindWordStartY(page, "Delivery");
        double rightHeadingY = FindWordStartY(page, "Narrative");
        double riskHeaderY = FindWordStartY(page, "Area");
        double riskBodyY = FindWordStartY(page, "PDF");
        double decisionHeaderY = FindWordStartY(page, "Next");
        double decisionBodyY = FindWordStartY(page, "fixtures");

        Assert.Equal(1, pdf.NumberOfPages);
        Assert.True(titleY - leadY >= 21, $"Expected dashboard lead copy to sit comfortably below the title. Gap: {titleY - leadY:0.##}pt.");
        Assert.True(leadY - firstMetricY >= 24, $"Expected metric cards to start after the lead copy with visible breathing room. Gap: {leadY - firstMetricY:0.##}pt.");
        Assert.True(Math.Abs(firstMetricY - secondMetricY) <= 0.5, "Expected metric card values to align on the same visual baseline.");
        Assert.True(Math.Abs(firstMetricY - thirdMetricY) <= 0.5, "Expected metric card values to align on the same visual baseline.");
        Assert.True(firstMetricY - fourthMetricY >= 10, "Expected the long fourth metric label to wrap below its value instead of colliding with neighboring cards.");
        Assert.True(leftHeadingY - riskHeaderY >= 175, $"Expected the trend drawing to reserve vertical space before the risk table. Gap: {leftHeadingY - riskHeaderY:0.##}pt.");
        Assert.True(Math.Abs(leftHeadingY - rightHeadingY) <= 2, "Expected the two body columns to start on the same visual row.");
        Assert.True(riskHeaderY - riskBodyY >= 14, $"Expected dashboard table header and first row to retain readable rhythm. Gap: {riskHeaderY - riskBodyY:0.##}pt.");
        Assert.True(decisionHeaderY - decisionBodyY >= 14, $"Expected decision table header and first row to retain readable rhythm. Gap: {decisionHeaderY - decisionBodyY:0.##}pt.");

        Assert.Contains(leftLines, line => line.Text.Contains("Deliverytrend", StringComparison.Ordinal));
        Assert.Contains(rightLines, line => line.Text.Contains("Narrative", StringComparison.Ordinal));
        Assert.True(FindWordStartX(page, "Narrative") >= rightColumnX - 1,
            $"Expected the right dashboard column to start after the gutter. Narrative x: {FindWordStartX(page, "Narrative"):0.##}, expected at least {rightColumnX:0.##}.");
        Assert.True(FindWordEndX(page, "Roadmap") <= leftColumnRightX + 1,
            "Expected the left risk table owner column to stay inside the left dashboard column.");
        Assert.True(FindWordEndX(page, "slices") <= rightColumnRightX + 1,
            "Expected the right decision table text to stay inside the right dashboard column.");

        AssertNoCrampedBaselines(leftLines, "showcase dashboard left column");
        AssertNoCrampedBaselines(rightLines, "showcase dashboard right column");
        AssertNoSameBaselineTextCollisions(page, "showcase dashboard");
        AssertNoAmbiguousSameBaselineRunGaps(page, "showcase dashboard");
    }


}
