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
    private static byte[] CreateTableCellSpacingProbe(PdfCreateOptions options, double cellSpacing, bool useRowColumnFlow) {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(1);
        style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(1);
        style.CellSpacing = global::OfficeIMO.OfficeLength.Points(cellSpacing);
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(90), global::OfficeIMO.OfficeLength.Points(90) };

        var rows = new[] {
            new[] { "SpacingA1", "SpacingB1" },
            new[] { "SpacingA2", "SpacingB2" }
        };

        if (useRowColumnFlow) {
            return PdfDocument.Create(options)
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.PercentColumn(100, column => column.Table(rows, style: style))))))
                .ToBytes();
        }

        return PdfDocument.Create(options)
            .Table(rows, style: style)
            .ToBytes();
    }

    private static byte[] CreateTablePaddingProbe(PdfCreateOptions options, bool useRowColumnFlow, bool useSidePadding) {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0);
        style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0);
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(90) };
        if (useSidePadding) {
            style.CellPaddingLeft = global::OfficeIMO.OfficeLength.Points(18);
            style.CellPaddingRight = global::OfficeIMO.OfficeLength.Points(3);
            style.CellPaddingTop = global::OfficeIMO.OfficeLength.Points(16);
            style.CellPaddingBottom = global::OfficeIMO.OfficeLength.Points(4);
        }

        var rows = new[] {
            new[] { "PadMarker" }
        };

        if (useRowColumnFlow) {
            return PdfDocument.Create(options)
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.PercentColumn(100, column => column.Table(rows, style: style))))))
                .ToBytes();
        }

        return PdfDocument.Create(options)
            .Table(rows, style: style)
            .ToBytes();
    }

    private static byte[] CreateTablePerCellPaddingProbe(PdfCreateOptions options, bool useRowColumnFlow, bool useCellPadding) {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0);
        style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0);
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(110) };
        if (useCellPadding) {
            style.CellPaddings = new Dictionary<(int Row, int Column), PdfCellPadding> {
                [(0, 0)] = new PdfCellPadding {
                    Left = global::OfficeIMO.OfficeLength.Points(22),
                    Right = global::OfficeIMO.OfficeLength.Points(3),
                    Top = global::OfficeIMO.OfficeLength.Points(16),
                    Bottom = global::OfficeIMO.OfficeLength.Points(4)
                }
            };
        }

        var rows = new[] {
            new[] { "CellPadMarker" }
        };

        if (useRowColumnFlow) {
            return PdfDocument.Create(options)
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.PercentColumn(100, column => column.Table(rows, style: style))))))
                .ToBytes();
        }

        return PdfDocument.Create(options)
            .Table(rows, style: style)
            .ToBytes();
    }

    private static byte[] CreateTablePerCellAlignmentProbe(PdfCreateOptions options, bool useRowColumnFlow, bool useCellAlignment) {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(2);
        style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(2);
        style.MinRowHeight = global::OfficeIMO.OfficeLength.Points(72);
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(130) };
        if (useCellAlignment) {
            style.CellAlignments = new Dictionary<(int Row, int Column), PdfColumnAlign> {
                [(0, 0)] = PdfColumnAlign.Right
            };
            style.CellVerticalAlignments = new Dictionary<(int Row, int Column), PdfCellVerticalAlign> {
                [(0, 0)] = PdfCellVerticalAlign.Bottom
            };
        }

        var rows = new[] {
            new[] { "CellAlignMarker" }
        };

        if (useRowColumnFlow) {
            return PdfDocument.Create(options)
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.PercentColumn(100, column => column.Table(rows, style: style))))))
                .ToBytes();
        }

        return PdfDocument.Create(options)
            .Table(rows, style: style)
            .ToBytes();
    }

    private static byte[] CreateTableLineHeightProbe(PdfCreateOptions options, double? lineHeight, bool useRowColumnFlow) {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.FontSize = global::OfficeIMO.OfficeLength.Points(10);
        style.LineHeight = lineHeight;
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(72) };

        var rows = new[] {
            new[] { "FirstLine SecondLine" }
        };

        if (useRowColumnFlow) {
            return PdfDocument.Create(options)
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.PercentColumn(100, column => column.Table(rows, style: style))))))
                .ToBytes();
        }

        return PdfDocument.Create(options)
            .Table(rows, style: style)
            .ToBytes();
    }

    private static byte[] CreateTableSpacingProbe(PdfCreateOptions options, double spacingBefore, double spacingAfter) {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore);
        style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(spacingAfter);

        return PdfDocument.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"))
            .Table(new[] {
                new[] { "Alpha", "Ready" },
                new[] { "Beta", "Ready" }
            }, style: style)
            .Paragraph(p => p.Text("AfterMarker"))
            .ToBytes();
    }

    private static byte[] CreateLightTableRhythmProbe(PdfCreateOptions options, PdfTableStyle? style) {
        return PdfDocument.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"))
            .Table(new[] {
                new[] { "Alpha", "Ready" },
                new[] { "Beta", "Ready" }
            }, style: style)
            .Paragraph(p => p.Text("AfterMarker"))
            .ToBytes();
    }

    private static byte[] CreateCompressionProbe(bool compressContentStreams) {
        PdfDocument doc = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(420),
            PageHeight = global::OfficeIMO.OfficeLength.Points(1200),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(42),
            MarginRight = global::OfficeIMO.OfficeLength.Points(42),
            MarginTop = global::OfficeIMO.OfficeLength.Points(42),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(42),
            CompressContentStreams = compressContentStreams
        })
            .H1("CompressionProbe");

        for (int i = 0; i < 18; i++) {
            doc.Paragraph(p => p.Text("CompressionProbe repeated body repeated body repeated body repeated body repeated body " + i.ToString(CultureInfo.InvariantCulture)));
        }

        return doc.ToBytes();
    }

    private static byte[] CreateMinimalIccProfile(string colorSpace = "RGB ") {
        byte[] profile = new byte[132];
        profile[0] = 0;
        profile[1] = 0;
        profile[2] = 0;
        profile[3] = 132;
        Encoding.ASCII.GetBytes(colorSpace, 0, 4, profile, 16);
        profile[36] = (byte)'a';
        profile[37] = (byte)'c';
        profile[38] = (byte)'s';
        profile[39] = (byte)'p';
        return profile;
    }

    private static string? FindLocalTrueTypeFont() {
        string windowsFont = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (File.Exists(windowsFont)) {
            return windowsFont;
        }

        string[] candidates = {
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
            "/Library/Fonts/Arial.ttf"
        };
        foreach (string candidate in candidates) {
            if (File.Exists(candidate)) {
                return candidate;
            }
        }

        return null;
    }

    private static string RenderTableStyleContent(PdfTableStyle style) {
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" }
            }, style: style)
            .ToBytes();

        return PdfTestContentText.ReadDecoded(bytes);
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int startIndex = 0;
        while (true) {
            int index = text.IndexOf(value, startIndex, StringComparison.Ordinal);
            if (index < 0) {
                return count;
            }

            count++;
            startIndex = index + value.Length;
        }
    }

    private static List<(double X1, double Y1, double X2, double Y2)> ExtractLinkRectangles(string pdf) {
        var rectangles = new List<(double X1, double Y1, double X2, double Y2)>();
        var matches = Regex.Matches(pdf, @"/Subtype /Link\b.*?/Rect \[(?<x1>-?\d+(?:\.\d+)?) (?<y1>-?\d+(?:\.\d+)?) (?<x2>-?\d+(?:\.\d+)?) (?<y2>-?\d+(?:\.\d+)?)\]");
        foreach (Match match in matches) {
            rectangles.Add((
                double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)));
        }

        return rectangles;
    }

    private static List<(double X, double Y, double W, double H)> ExtractPaintedRectangles(string content, string colorOperator, string paintOperator) {
        var rectangles = new List<(double X, double Y, double W, double H)>();
        string pattern = Regex.Escape(colorOperator) +
            @"[\s\S]*?(?<x>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) (?<w>-?\d+(?:\.\d+)?) (?<h>-?\d+(?:\.\d+)?) re\s+" +
            Regex.Escape(paintOperator);
        var matches = Regex.Matches(content, pattern);
        foreach (Match match in matches) {
            rectangles.Add((
                double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["h"].Value, CultureInfo.InvariantCulture)));
        }

        return rectangles;
    }

    private static List<(double X1, double Y1, double X2, double Y2)> ExtractStrokedLineSegments(string content, string colorOperator) {
        var segments = new List<(double X1, double Y1, double X2, double Y2)>();
        const string number = @"-?\d+(?:\.\d+)?";
        string pattern = Regex.Escape(colorOperator) +
            @"\s+(?:" + number + @" w\s+)?" +
            @"(?<x1>" + number + @") (?<y1>" + number + @") m\s+" +
            @"(?<x2>" + number + @") (?<y2>" + number + @") l\s+S";
        var matches = Regex.Matches(content, pattern);
        foreach (Match match in matches) {
            segments.Add((
                double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)));
        }

        return segments;
    }

    private static byte[] CreateParagraphSpacingProbe(PdfCreateOptions options, PdfParagraphStyle? style) {
        return PdfDocument.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"))
            .Paragraph(p => p.Text("TargetMarker"), style: style)
            .Paragraph(p => p.Text("AfterMarker"))
            .ToBytes();
    }

    private static byte[] CreatePanelSpacingProbe(PdfCreateOptions options, PdfPanelStyle style) {
        return PdfDocument.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            })
            .PanelParagraph(p => p.Text("PanelMarker"), style)
            .Paragraph(p => p.Text("AfterMarker"))
            .ToBytes();
    }

    private static byte[] CreateParagraphLineHeightProbe(PdfCreateOptions options, PdfParagraphStyle? style) {
        return PdfDocument.Create(options)
            .Paragraph(p => p
                .Text("FirstLine")
                .LineBreak()
                .Text("SecondLine")
                .LineBreak()
                .Text("ThirdLine"), style: style)
            .ToBytes();
    }

    private static byte[] CreateParagraphIndentProbe(PdfCreateOptions options, PdfParagraphStyle? style) {
        return PdfDocument.Create(options)
            .Paragraph(p => p.Text("IndentedMarker alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau."), style: style)
            .ToBytes();
    }

    private static byte[] CreateTopLevelFlowSpacingBeforeProbe(string blockKind, double spacingBefore) {
        var options = CreateFlowSpacingProbeOptions();
        var paragraphStyle = new PdfParagraphStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) };
        var doc = PdfDocument.Create(options);

        switch (blockKind) {
            case "bullet-list":
                doc.Bullets(new[] { "ListTopMarker" }, style: new PdfListStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), ItemSpacing = global::OfficeIMO.OfficeLength.Points(0) });
                break;
            case "numbered-list":
                doc.Numbered(new[] { "ListTopMarker" }, style: new PdfListStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), ItemSpacing = global::OfficeIMO.OfficeLength.Points(0) });
                break;
            case "panel":
                doc.PanelParagraph(p => p.Text("PanelTopMarker"), new PdfPanelStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), PaddingX = global::OfficeIMO.OfficeLength.Points(4), PaddingY = global::OfficeIMO.OfficeLength.Points(4) });
                break;
            case "horizontal-rule":
                doc.HR(style: new PdfHorizontalRuleStyle { Thickness = global::OfficeIMO.OfficeLength.Points(2), SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                    .Paragraph(p => p.Text("AfterFixedMarker"), style: paragraphStyle);
                break;
            case "image":
                doc.Image(CreateMinimalRgbPng(), 24, 12, style: new PdfImageStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                    .Paragraph(p => p.Text("AfterFixedMarker"), style: paragraphStyle);
                break;
            case "shape":
                doc.Shape(OfficeShape.Rectangle(24, 12), style: new PdfDrawingStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                    .Paragraph(p => p.Text("AfterFixedMarker"), style: paragraphStyle);
                break;
            case "drawing":
                doc.Drawing(new OfficeDrawing(24, 12).AddShape(OfficeShape.Rectangle(24, 12), 0, 0), style: new PdfDrawingStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                    .Paragraph(p => p.Text("AfterFixedMarker"), style: paragraphStyle);
                break;
            case "row":
                doc.Compose(document => document.Page(page => page.Content(content => content.Row(row => row
                    .Style(new PdfRowStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                    .PercentColumn(100, column => column.Paragraph(p => p.Text("RowTopMarker"), style: paragraphStyle))))));
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(blockKind), blockKind, "Unknown flow block kind.");
        }

        return doc.ToBytes();
    }

    private static byte[] CreateColumnFlowSpacingBeforeProbe(string blockKind, double spacingBefore) {
        var options = CreateFlowSpacingProbeOptions();
        return PdfDocument.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row
                .PercentColumn(100, column => AddColumnFlowSpacingBeforeProbe(column, blockKind, spacingBefore))))))
            .ToBytes();
    }

    private static void AddColumnFlowSpacingBeforeProbe(PdfContentBuilder column, string blockKind, double spacingBefore) {
        var paragraphStyle = new PdfParagraphStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) };

        switch (blockKind) {
            case "bullet-list":
                column.Bullets(new[] { "ColumnListMarker" }, style: new PdfListStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), ItemSpacing = global::OfficeIMO.OfficeLength.Points(0) });
                break;
            case "numbered-list":
                column.Numbered(new[] { "ColumnListMarker" }, style: new PdfListStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), ItemSpacing = global::OfficeIMO.OfficeLength.Points(0) });
                break;
            case "panel":
                column.PanelParagraph(p => p.Text("ColumnPanelMarker"), new PdfPanelStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), PaddingX = global::OfficeIMO.OfficeLength.Points(4), PaddingY = global::OfficeIMO.OfficeLength.Points(4) });
                break;
            case "horizontal-rule":
                column.HR(style: new PdfHorizontalRuleStyle { Thickness = global::OfficeIMO.OfficeLength.Points(2), SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                    .Paragraph(p => p.Text("ColumnAfterFixedMarker"), style: paragraphStyle);
                break;
            case "image":
                column.Image(CreateMinimalRgbPng(), global::OfficeIMO.OfficeLength.Points(24), global::OfficeIMO.OfficeLength.Points(12), style: new PdfImageStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                    .Paragraph(p => p.Text("ColumnAfterFixedMarker"), style: paragraphStyle);
                break;
            case "shape":
                column.Shape(OfficeShape.Rectangle(24, 12), style: new PdfDrawingStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                    .Paragraph(p => p.Text("ColumnAfterFixedMarker"), style: paragraphStyle);
                break;
            case "drawing":
                column.Drawing(new OfficeDrawing(24, 12).AddShape(OfficeShape.Rectangle(24, 12), 0, 0), style: new PdfDrawingStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                    .Paragraph(p => p.Text("ColumnAfterFixedMarker"), style: paragraphStyle);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(blockKind), blockKind, "Unknown flow block kind.");
        }
    }

    private static PdfCreateOptions CreateFlowSpacingProbeOptions() {
        return new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(320),
            PageHeight = global::OfficeIMO.OfficeLength.Points(220),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };
    }

    private static OfficeDrawing CreateKeepWithNextDrawingScene() {
        var shape = OfficeShape.Rectangle(24, 24);
        shape.FillColor = OfficeColor.WhiteSmoke;

        return new OfficeDrawing(24, 24)
            .AddShape(shape, 0, 0);
    }

    private static byte[] CreateMinimalRgbPng() => PdfPngTestImages.CreateRgbPng(255, 0, 0);

    private static double FindWordStartX(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.X;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static double FindWordEndX(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index + word.Length - 1].EndBaseLine.X;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static void AssertStatementRowColumns(UglyToad.PdfPig.Content.Page page, string productWord, string unitPrice, string quantity, string total, string rowName) {
        const double minimumProductGap = 10;
        const double minimumNumericGap = 6;
        double baselineY = FindWordStartY(page, productWord);
        var ordered = GetLettersOnBaseline(page, baselineY);
        string text = string.Concat(ordered.Select(letter => letter.Value));
        int productIndex = text.IndexOf(productWord, StringComparison.Ordinal);
        int unitPriceIndex = productIndex >= 0 ? text.IndexOf(unitPrice, productIndex + productWord.Length, StringComparison.Ordinal) : -1;
        int quantityIndex = unitPriceIndex >= 0 ? text.IndexOf(quantity, unitPriceIndex + unitPrice.Length, StringComparison.Ordinal) : -1;
        int totalIndex = quantityIndex >= 0 ? text.IndexOf(total, quantityIndex + quantity.Length, StringComparison.Ordinal) : -1;

        Assert.True(productIndex >= 0 && unitPriceIndex >= 0 && quantityIndex >= 0 && totalIndex >= 0,
            $"Expected {rowName} to contain product, unit price, quantity, and total tokens in left-to-right order. Text: '{text}'.");
        double productEndX = ordered[productIndex + productWord.Length - 1].EndBaseLine.X;
        double unitPriceStartX = ordered[unitPriceIndex].StartBaseLine.X;
        double unitPriceEndX = ordered[unitPriceIndex + unitPrice.Length - 1].EndBaseLine.X;
        double quantityStartX = ordered[quantityIndex].StartBaseLine.X;
        double quantityEndX = ordered[quantityIndex + quantity.Length - 1].EndBaseLine.X;
        double totalStartX = ordered[totalIndex].StartBaseLine.X;

        Assert.True(productEndX < unitPriceStartX - minimumProductGap,
            $"Expected {rowName} product text to end with visible space before the unit price column.");
        Assert.True(unitPriceEndX < quantityStartX - minimumNumericGap,
            $"Expected {rowName} unit price text to stay separated from the quantity column.");
        Assert.True(quantityEndX < totalStartX - minimumNumericGap,
            $"Expected {rowName} quantity text to stay separated from the total column.");
    }

    private static double FindWordStartXOnBaseline(UglyToad.PdfPig.Content.Page page, string word, double baselineY) {
        var ordered = GetLettersOnBaseline(page, baselineY);
        string text = string.Concat(ordered.Select(letter => letter.Value));
        int index = text.IndexOf(word, StringComparison.Ordinal);
        if (index >= 0) {
            return ordered[index].StartBaseLine.X;
        }

        throw new InvalidOperationException("Could not find word '" + word + "' on baseline " + baselineY.ToString("0.##", CultureInfo.InvariantCulture) + ".");
    }

    private static double FindWordEndXOnBaseline(UglyToad.PdfPig.Content.Page page, string word, double baselineY) {
        var ordered = GetLettersOnBaseline(page, baselineY);
        string text = string.Concat(ordered.Select(letter => letter.Value));
        int index = text.IndexOf(word, StringComparison.Ordinal);
        if (index >= 0) {
            return ordered[index + word.Length - 1].EndBaseLine.X;
        }

        throw new InvalidOperationException("Could not find word '" + word + "' on baseline " + baselineY.ToString("0.##", CultureInfo.InvariantCulture) + ".");
    }

    private static List<UglyToad.PdfPig.Content.Letter> GetLettersOnBaseline(UglyToad.PdfPig.Content.Page page, double baselineY) {
        double roundedBaselineY = Math.Round(baselineY, 1);
        return page.Letters
            .Where(letter =>
                !string.IsNullOrWhiteSpace(letter.Value) &&
                Math.Round(letter.StartBaseLine.Y, 1).Equals(roundedBaselineY))
            .OrderBy(letter => letter.StartBaseLine.X)
            .ToList();
    }

    private static double FindWordStartY(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.Y;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static double AverageBaselineY(IReadOnlyList<UglyToad.PdfPig.Content.Letter> letters, string word) {
        string text = string.Concat(letters.Select(letter => letter.Value));
        int index = text.IndexOf(word, StringComparison.Ordinal);
        if (index < 0) {
            throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
        }

        return letters
            .Skip(index)
            .Take(word.Length)
            .Average(letter => letter.StartBaseLine.Y);
    }

    private static double FindWordPointSize(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].PointSize;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static int CountTextLines(UglyToad.PdfPig.Content.Page page) {
        return page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
    }

    private static List<List<UglyToad.PdfPig.Content.Letter>> GetNonWhitespaceLetterLines(UglyToad.PdfPig.Content.Page page) {
        return page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderByDescending(group => group.Key)
            .Select(group => group.OrderBy(letter => letter.StartBaseLine.X).ToList())
            .ToList();
    }

    private static List<VisualTextLine> GetVisualTextLines(UglyToad.PdfPig.Content.Page page, double minX, double maxX) {
        return page.Letters
            .Where(letter =>
                !string.IsNullOrWhiteSpace(letter.Value) &&
                letter.StartBaseLine.X >= minX &&
                letter.EndBaseLine.X <= maxX)
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderByDescending(group => group.Key)
            .Select(group => {
                var letters = group.OrderBy(letter => letter.StartBaseLine.X).ToList();
                return new VisualTextLine(
                    string.Concat(letters.Select(letter => letter.Value)),
                    letters.Min(letter => letter.StartBaseLine.X),
                    letters.Max(letter => letter.EndBaseLine.X),
                    letters[0].StartBaseLine.Y);
            })
            .ToList();
    }

    private static void AssertReadableTextRhythm(IReadOnlyList<VisualTextLine> lines, string areaName) {
        for (int i = 1; i < lines.Count; i++) {
            double baselineGap = lines[i - 1].BaselineY - lines[i].BaselineY;
            Assert.True(baselineGap >= 9.5,
                $"Expected {areaName} text to keep readable baseline rhythm between '{lines[i - 1].Text}' and '{lines[i].Text}'. Gap: {baselineGap:0.##}pt.");
        }
    }

    private static void AssertNoCrampedBaselines(IReadOnlyList<VisualTextLine> lines, string areaName) {
        for (int i = 1; i < lines.Count; i++) {
            double baselineGap = lines[i - 1].BaselineY - lines[i].BaselineY;
            Assert.True(baselineGap >= 7.5,
                $"Expected {areaName} to avoid cramped or overlapping baselines between '{lines[i - 1].Text}' and '{lines[i].Text}'. Gap: {baselineGap:0.##}pt.");
        }
    }

    private static void AssertNoSameBaselineTextCollisions(UglyToad.PdfPig.Content.Page page, string areaName) {
        foreach (var line in page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            for (int i = 1; i < ordered.Count; i++) {
                double gap = ordered[i].StartBaseLine.X - ordered[i - 1].EndBaseLine.X;
                // Pair kerning (AFM and GPOS pairs stay within about 0.2 em) overlaps advance boxes by design.
                double kerningTolerance = Math.Max(0.25, 0.2 * Math.Min(ordered[i - 1].PointSize, ordered[i].PointSize));
                Assert.True(gap >= -kerningTolerance,
                    $"Expected {areaName} text not to collide on baseline {line.Key:0.##}, but '{ordered[i - 1].Value}' and '{ordered[i].Value}' overlap by {-gap:0.##}pt.");
            }
        }
    }

    private static void AssertNoAmbiguousSameBaselineRunGaps(UglyToad.PdfPig.Content.Page page, string areaName) {
        const double sameRunGap = 0.75;
        const double minimumReadableRunGap = 1.8;

        foreach (var line in page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            for (int i = 1; i < ordered.Count; i++) {
                double gap = ordered[i].StartBaseLine.X - ordered[i - 1].EndBaseLine.X;
                if (gap <= sameRunGap) {
                    continue;
                }

                Assert.True(gap >= minimumReadableRunGap,
                    $"Expected {areaName} to avoid visually ambiguous same-baseline run gaps on baseline {line.Key:0.##} between '{ordered[i - 1].Value}' and '{ordered[i].Value}'. Gap: {gap:0.##}pt.");
            }
        }
    }

    private static List<double> GetInterWordGaps(List<UglyToad.PdfPig.Content.Letter> letters) {
        return letters
            .Zip(letters.Skip(1), (left, right) => right.StartBaseLine.X - left.EndBaseLine.X)
            .Where(gap => gap > 1)
            .ToList();
    }

    private static IReadOnlyList<string> GetPageContentStreams(byte[] pdf, int pageNumber) {
        var document = PdfReadDocument.Open(pdf);
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        int pageObjectNumber = document.Pages[pageNumber - 1].ObjectNumber;
        if (!objects.TryGetValue(pageObjectNumber, out var pageObject) || pageObject.Value is not PdfDictionary pageDictionary) {
            throw new InvalidOperationException("Page object was not found.");
        }

        if (!pageDictionary.Items.TryGetValue("Contents", out var contents)) {
            throw new InvalidOperationException("Page contents were not found.");
        }

        var streams = new List<string>();
        AppendContentStreams(objects, contents, streams);
        return streams;
    }

    private static void AppendContentStreams(Dictionary<int, PdfIndirectObject> objects, PdfObject contents, List<string> streams) {
        if (contents is PdfReference reference) {
            if (objects.TryGetValue(reference.ObjectNumber, out var indirect) && indirect.Value is PdfStream stream) {
                streams.Add(Encoding.GetEncoding("ISO-8859-1").GetString(OfficeIMO.Pdf.Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects)));
            }

            return;
        }

        if (contents is PdfArray array) {
            foreach (var item in array.Items) {
                AppendContentStreams(objects, item, streams);
            }
        }
    }

    private sealed class VisualTextLine {
        internal VisualTextLine(string text, double x1, double x2, double baselineY) {
            Text = text;
            X1 = x1;
            X2 = x2;
            BaselineY = baselineY;
        }

        internal string Text { get; }

        internal double X1 { get; }

        internal double X2 { get; }

        internal double BaselineY { get; }
    }

}
