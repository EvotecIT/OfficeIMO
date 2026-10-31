using W = DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Core;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("header")]
    [InlineData("conditional")]
    [InlineData("merged-header")]
    [InlineData("repeated-header")]
    [InlineData("repeated-merged-header")]
    [InlineData("top")]
    [InlineData("bottom")]
    [InlineData("left")]
    [InlineData("right")]
    public void CollapsedCellBorderOwnsItsSharedEdgeInsteadOfTheTableBorder(string side) {
        using var document = WordDocument.Create();
        bool header = side.EndsWith("header", StringComparison.Ordinal);
        bool repeated = side.StartsWith("repeated-", StringComparison.Ordinal);
        var table = document.AddTable(repeated ? 16 : 3, 2, header ? WordTableStyle.OfficeIMOModern : WordTableStyle.TableGrid);
        if (header) {
            if (side == "merged-header") table.MergeCells(1, 0, 1, 2);
        } else if (side == "conditional") {
            const string id = "ConditionalVisibleGrid";
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new W.Style(
                new W.StyleName { Val = id },
                new W.StyleTableProperties(new W.TableBorders(
                    new W.TopBorder { Val = W.BorderValues.Single, Color = "808080", Size = 16U },
                    new W.BottomBorder { Val = W.BorderValues.Single, Color = "808080", Size = 16U },
                    new W.LeftBorder { Val = W.BorderValues.Single, Color = "808080", Size = 16U },
                    new W.RightBorder { Val = W.BorderValues.Single, Color = "808080", Size = 16U },
                    new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Color = "808080", Size = 16U },
                    new W.InsideVerticalBorder { Val = W.BorderValues.Single, Color = "808080", Size = 16U })),
                new W.TableStyleProperties(new W.TableStyleConditionalFormattingTableCellProperties(new W.TableCellBorders(
                    new W.BottomBorder { Val = W.BorderValues.Single, Color = "FF0000", Size = 8U }))) { Type = W.TableStyleOverrideValues.FirstRow })
                { Type = W.StyleValues.Table, StyleId = id });
            table._tableProperties!.TableStyle = new W.TableStyle { Val = id };
            table.ConditionalFormattingFirstRow = true;
        } else {
            table.ConditionalFormattingFirstRow = false;
            table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Single, OfficeLength.Points((16U) / 8d), OfficeColor.Gray);
            var border = table.Rows[side == "bottom" ? 0 : 1].Cells[side == "left" ? 1 : 0].Borders;
            switch (side) {
                case "top": border.TopStyle = WordBorderStyle.Single; border.TopColorHex = "FF0000"; border.TopSize = OfficeLength.Points((8U) / 8d); break;
                case "bottom": border.BottomStyle = WordBorderStyle.Single; border.BottomColorHex = "FF0000"; border.BottomSize = OfficeLength.Points((8U) / 8d); break;
                case "left": border.LeftStyle = WordBorderStyle.Single; border.LeftColorHex = "FF0000"; border.LeftSize = OfficeLength.Points((8U) / 8d); break;
                case "right": border.RightStyle = WordBorderStyle.Single; border.RightColorHex = "FF0000"; border.RightSize = OfficeLength.Points((8U) / 8d); break;
            }
        }
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Header";
        table.Rows[1].Cells[0].Paragraphs[0].Text = "Body";
        if (side == "left" || side == "right") {
            var target = table.Rows[1].Cells[side == "right" ? 1 : 0];
            target.MarginLeft = target.MarginRight = OfficeLength.Zero;
            target.Paragraphs[0].Text = "Edge";
            if (side == "left") target.Paragraphs[0].ParagraphAlignment = WordParagraphAlignment.Right;
        }
        for (int row = 2; row < table.Rows.Count; row++) table.Rows[row].Cells[0].Paragraphs[0].Text = $"Body {row}";
        if (repeated) {
            table.Rows[1].Cells[0].Borders.BottomStyle = WordBorderStyle.None;
            table.Rows[1].Cells[0].Paragraphs[0].Text = "BeforeBoundary";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "AfterBoundary";
        }
        if (side == "repeated-merged-header") table.MergeCells(3, 0, 4, 1);
        var options = new WordPdfExportOptions { IncludePageNumbers = false };
        if (repeated) {
            options.PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeLength.Points(300), OfficeLength.Points(180));
            options.Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(OfficeLength.Points(20));
        }
        using var converted = document.ToPdf(options);
        using var pdf = PdfPigDocument.Open(converted.ToBytes());
        if (repeated) Assert.True(pdf.NumberOfPages > 1);
        foreach (var page in pdf.GetPages()) {
            var strokes = page.Paths.Where(path => path.IsStroked)
                .SelectMany(path => path.SelectMany(SharedEdgeSegments)
                    .Select(line => (Color: path.StrokeColor.ToRGBValues(), Box: line.GetBoundingRectangle())))
                .ToArray();
            if (repeated && page.Number == 1) {
                var before = Assert.Single(page.GetWords(), word => word.Text == "BeforeBoundary").BoundingBox;
                var after = Assert.Single(page.GetWords(), word => word.Text == "AfterBoundary").BoundingBox;
                Assert.Contains(strokes, stroke => stroke.Box.Height < .01D &&
                    stroke.Box.Top > after.Top && stroke.Box.Top < before.Bottom &&
                    stroke.Box.Left < after.Left + 1D && stroke.Box.Right > after.Left + 1D);
            }
            var color = header ? (0.247D, 0.392D, 0.843D) : (1D, 0D, 0D);
            var winners = strokes.Where(stroke => stroke.Color == color).ToArray();
            Assert.NotEmpty(winners);
            foreach (var winner in winners) {
                Assert.DoesNotContain(strokes, other => other.Color != color && SharedLineOverlap(winner.Box, other.Box));
            }
            if (side == "left" || side == "right") {
                var edge = Assert.Single(page.GetWords(), word => word.Text == "Edge");
                double boundary = winners[0].Box.Left;
                if (side == "right") Assert.InRange(edge.Letters[0].StartBaseLine.X - boundary, .49D, double.PositiveInfinity);
                else Assert.InRange(boundary - edge.Letters[edge.Letters.Count - 1].EndBaseLine.X, .49D, double.PositiveInfinity);
            }
        }
    }

    [Theory]
    [InlineData("header-nil")]
    [InlineData("body-cell")]
    [InlineData("no-repeat-top")]
    [InlineData("no-repeat-bottom")]
    public void CollapsedTableBordersResolveAgainstNeighboursOnThePhysicalPage(string profile) {
        using var document = WordDocument.Create();
        var table = document.AddTable(16, 1, WordTableStyle.TableGrid);
        table.ConditionalFormattingFirstRow = false;
        table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Single, OfficeLength.Points((16U) / 8d), OfficeColor.Gray);
        table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = profile == "header-nil" || profile == "body-cell";
        for (int row = 0; row < table.Rows.Count; row++) {
            var cell = table.Rows[row].Cells[0];
            cell.Paragraphs[0].Text = row == 0 ? "Header" : "Body" + row;
            if (profile == "no-repeat-top") {
                cell.Borders.BottomStyle = WordBorderStyle.Single; cell.Borders.BottomColorHex = "FF0000"; cell.Borders.BottomSize = OfficeLength.Points((8U) / 8d);
            } else if ((profile == "body-cell" && row > 1) || (profile == "no-repeat-bottom" && row > 0)) {
                cell.Borders.TopStyle = WordBorderStyle.Single; cell.Borders.TopColorHex = "FF0000"; cell.Borders.TopSize = OfficeLength.Points((8U) / 8d);
            }
        }
        if (profile == "header-nil") table.Rows[0].Cells[0].Borders.BottomStyle = WordBorderStyle.Nil;
        using var converted = document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeLength.Points(300), OfficeLength.Points(180)),
            Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(OfficeLength.Points(20))
        });
        using var pdf = PdfPigDocument.Open(converted.ToBytes());
        Assert.True(pdf.NumberOfPages > 1);
        foreach (var page in pdf.GetPages()) {
            var strokes = page.Paths.Where(path => path.IsStroked)
                .SelectMany(path => path.SelectMany(SharedEdgeSegments)
                    .Select(line => (Color: path.StrokeColor.ToRGBValues(), Box: line.GetBoundingRectangle())))
                .ToArray();
            if (profile == "body-cell") {
                var winners = strokes.Where(stroke => stroke.Color == (1D, 0D, 0D)).ToArray();
                Assert.NotEmpty(winners);
                foreach (var winner in winners)
                    Assert.DoesNotContain(strokes, other => other.Color != winner.Color && SharedLineOverlap(winner.Box, other.Box));
            } else if (profile == "header-nil") {
                var header = Assert.Single(page.GetWords(), word => word.Text == "Header").BoundingBox;
                var body = page.GetWords().Where(word => word.Text.StartsWith("Body", StringComparison.Ordinal))
                    .OrderByDescending(word => word.BoundingBox.Top).First().BoundingBox;
                Assert.DoesNotContain(strokes, stroke => stroke.Box.Height < .01D && stroke.Box.Top > body.Top && stroke.Box.Top < header.Bottom);
            } else if ((profile == "no-repeat-top" && page.Number > 1) || (profile == "no-repeat-bottom" && page.Number < pdf.NumberOfPages)) {
                var vertical = strokes.Where(stroke => stroke.Box.Width < .01D && stroke.Box.Height > .01D).ToArray();
                double boundary = profile == "no-repeat-top" ? vertical.Max(stroke => stroke.Box.Top) : vertical.Min(stroke => stroke.Box.Bottom);
                Assert.Contains(strokes, stroke => stroke.Box.Height < .01D && Math.Abs(stroke.Box.Top - boundary) < .01D && stroke.Box.Width > .01D);
            }
        }
    }

    private static IEnumerable<PdfLine> SharedEdgeSegments(PdfSubpath path) {
        PdfPoint? start = null, previous = null;
        foreach (var command in path.Commands) {
            if (command is PdfSubpath.Move move) start = previous = move.Location;
            else if (command is PdfSubpath.Line line) {
                yield return new PdfLine(line.From, line.To);
                previous = line.To;
            } else if (command is PdfSubpath.Close && previous.HasValue && start.HasValue) {
                yield return new PdfLine(previous.Value, start.Value);
            }
        }
    }

    private static bool SharedLineOverlap(PdfRectangle first, PdfRectangle second) {
        const double tolerance = .01D;
        if (first.Height < tolerance && second.Height < tolerance && Math.Abs(first.Top - second.Top) < tolerance)
            return Math.Min(first.Right, second.Right) - Math.Max(first.Left, second.Left) > tolerance;
        if (first.Width < tolerance && second.Width < tolerance && Math.Abs(first.Left - second.Left) < tolerance)
            return Math.Min(first.Top, second.Top) - Math.Max(first.Bottom, second.Bottom) > tolerance;
        return false;
    }
}
