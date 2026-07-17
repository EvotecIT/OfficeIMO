using OfficeIMO.Drawing;
using System.Text;

namespace OfficeIMO.OneNote.Tests;

public sealed class OneNoteRenderingTests {
    [Fact]
    public void PageRendererProjectsRichTextImageInkMathTableAndAttachments() {
        OneNotePage page = CreateVisualPage();

        OneNotePageVisualSnapshot snapshot = OneNotePageRenderer.CreateSnapshot(page, new OneNotePageRenderingOptions {
            AutomaticPageWidthPoints = 360D,
            AutomaticPageHeightPoints = 420D
        });

        Assert.Equal(216D, snapshot.Drawing.Width);
        Assert.Equal(360D, snapshot.Drawing.Height);
        Assert.NotEmpty(snapshot.Drawing.Elements);
        Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingImage);
        Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingShape);
        Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText || element is OfficeDrawingRichText);
        Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
    }

    [Fact]
    public void PageExportsPngJpegTiffSvgAndWebpThroughSharedDrawingEncoders() {
        OneNotePage page = CreateVisualPage();

        OfficeImageExportResult png = page.ToImage().WithScale(0.25D).AsPng().Export();
        OfficeImageExportResult jpeg = page.ToImage().WithScale(0.25D).AsJpeg().Export();
        OfficeImageExportResult tiff = page.ToImage().WithScale(0.25D).AsTiff().Export();
        OfficeImageExportResult svg = page.ToImage().WithScale(0.25D).AsSvg().Export();
        OfficeImageExportResult webp = page.ToImage().WithScale(0.25D).AsWebp().Export();

        Assert.Equal(new byte[] { 137, 80, 78, 71 }, png.Bytes.Take(4));
        Assert.Equal(new byte[] { 0xFF, 0xD8 }, jpeg.Bytes.Take(2));
        Assert.Equal("II", Encoding.ASCII.GetString(tiff.Bytes, 0, 2));
        Assert.Contains("<svg", Encoding.UTF8.GetString(svg.Bytes), StringComparison.Ordinal);
        Assert.Equal("RIFF", Encoding.ASCII.GetString(webp.Bytes, 0, 4));
        Assert.Equal("WEBP", Encoding.ASCII.GetString(webp.Bytes, 8, 4));
        Assert.All(new[] { png, jpeg, tiff, svg, webp }, result => {
            Assert.True(result.Width > 0);
            Assert.True(result.Height > 0);
            Assert.True(result.Bytes.Length > 20);
        });
    }

    [Fact]
    public void SectionAndNotebookBatchExportUseStablePageSelectionAndNames() {
        var section = new OneNoteSection { Name = "Visuals" };
        section.Pages.Add(CreateVisualPage("First"));
        section.Pages.Add(CreateVisualPage("Second"));
        var notebook = new OneNoteNotebook { Name = "Notebook" };
        notebook.Sections.Add(section);

        IReadOnlyList<OfficeImageExportResult> sectionResults = section.ToImages().FromPage(1).TakePages(1).WithScale(0.2D).AsPng().Export();
        IReadOnlyList<OfficeImageExportResult> overflowSafeResults = section.ToImages().FromPage(1).TakePages(int.MaxValue).WithScale(0.2D).AsPng().Export();
        IReadOnlyList<OfficeImageExportResult> notebookResults = notebook.ToImages().AllPages().WithScale(0.2D).AsSvg().Export();

        Assert.Single(sectionResults);
        Assert.Single(overflowSafeResults);
        Assert.Equal("Second", sectionResults[0].Name);
        Assert.Equal("Second", overflowSafeResults[0].Name);
        Assert.Equal(2, notebookResults.Count);
        Assert.Equal(new[] { "First", "Second" }, notebookResults.Select(result => result.Name));
        Assert.All(notebookResults, result => Assert.Contains("Notebook/Visuals", result.Source, StringComparison.Ordinal));
    }

    [Fact]
    public void AutomaticCanvasExpandsPastStoredDimensionsToIncludeAbsoluteInk() {
        var page = new OneNotePage { Width = 8D, Height = 8D, PageSize = OneNotePageSize.Automatic };
        var ink = new OneNoteInk { Layout = new OneNoteLayout { X = 0D, Y = 0D } };
        ink.Ink.Add(new OfficeInkStroke().AddPoint(1D, 12D).AddPoint(2D, 13D));
        page.DirectContent.Add(ink);

        OfficeDrawing drawing = page.ToDrawing();

        Assert.True(drawing.Height >= 13D * OneNotePageRenderer.PointsPerHalfInch);
        Assert.True(drawing.Elements.Count(element => element is OfficeDrawingShape) > 1);
    }

    [Fact]
    public void FlowInkAndStandaloneMathStayWithinTheirAllocatedOutlineWidth() {
        var page = new OneNotePage { PageSize = OneNotePageSize.IndexCard };
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 0.25D, Y = 0.5D, Width = 1D } };
        var ink = new OneNoteInk();
        ink.Ink.Add(new OfficeInkStroke().AddPoint(0D, 0.2D).AddPoint(4D, 0.2D));
        outline.Children.Add(ink);
        outline.Children.Add(new OneNoteMath().SetExpression(OfficeMath.Identifier("MMMMMMMMMMMM")));
        page.Outlines.Add(outline);

        OneNotePageVisualSnapshot snapshot = OneNotePageRenderer.CreateSnapshot(page);
        OfficeDrawingShape line = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(),
            shape => shape.Shape.Kind == OfficeShapeKind.Line);
        OfficeDrawingText fallback = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>());
        double outlineLeft = 0.25D * OneNotePageRenderer.PointsPerHalfInch;
        double outlineRight = outlineLeft + OneNotePageRenderer.PointsPerHalfInch;

        Assert.True(line.X + line.Shape.Width <= outlineRight + 0.001D);
        Assert.True(fallback.X + fallback.Width <= outlineRight + 0.001D);
        Assert.Contains(snapshot.Diagnostics, diagnostic => diagnostic.Code == "ONENOTE_RENDER_MATH_CLIPPED");
    }

    [Fact]
    public void RasterExportLimitsLargeNativeCanvasesWithoutAllocatingTheRequestedSurface() {
        var page = new OneNotePage {
            PageSize = OneNotePageSize.Custom,
            Width = 100000D / OneNotePageRenderer.PointsPerHalfInch,
            Height = 100000D / OneNotePageRenderer.PointsPerHalfInch
        };
        var options = new OneNotePageRenderingOptions { Scale = 1D, MaximumRasterPixels = 10_000L };

        OfficeImageExportResult result = OneNotePageImageRenderer.Render(page, OfficeImageExportFormat.Png, options);

        Assert.Equal(100, result.Width);
        Assert.Equal(100, result.Height);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ONENOTE_IMAGE_RASTER_SCALE_LIMITED");
    }

    [Fact]
    public void RasterExportReportsAndPaintsUnsupportedSourceImagesInsteadOfDroppingThem() {
        var page = new OneNotePage { PageSize = OneNotePageSize.IndexCard };
        page.DirectContent.Add(new OneNoteImage {
            FileName = "vector.svg",
            MediaType = "image/svg+xml",
            Layout = new OneNoteLayout { X = 0.5, Y = 0.5, Width = 2, Height = 1 },
            Payload = OneNoteBinaryPayload.FromBytes(Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"20\" height=\"10\"><rect width=\"20\" height=\"10\" fill=\"red\"/></svg>"))
        });

        OfficeImageExportResult result = page.ToImage().WithScale(0.25D).AsPng().Export();
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DRAWING_RASTER_IMAGE_UNSUPPORTED");
        Assert.True(OfficeRasterImageDecoder.TryDecode(result.Bytes, out OfficeRasterImage? raster));
        Assert.NotNull(raster);
        Assert.Contains(raster!.GetPixels(), value => value != byte.MaxValue);
    }

    [Fact]
    public void SvgExportReportsAndEmbedsAVisibleFallbackForTiffSourceImages() {
        byte[] sourceTiff = OfficeRasterImageEncoder.Encode(
            new OfficeRasterImage(8, 4, OfficeColor.CornflowerBlue),
            OfficeImageExportFormat.Tiff);
        var page = new OneNotePage { PageSize = OneNotePageSize.IndexCard };
        page.DirectContent.Add(new OneNoteImage {
            FileName = "scan.tiff",
            MediaType = "image/tiff",
            Layout = new OneNoteLayout { X = 0.5, Y = 0.5, Width = 2, Height = 1 },
            Payload = OneNoteBinaryPayload.FromBytes(sourceTiff)
        });

        OfficeImageExportResult result = page.ToImage().WithScale(0.25D).AsSvg().Export();
        string svg = Encoding.UTF8.GetString(result.Bytes);

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DRAWING_RASTER_IMAGE_UNSUPPORTED");
        Assert.Contains("data:image/png;base64,", svg, StringComparison.Ordinal);
        Assert.Contains("<image", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void InlineMathParagraphKeepsListTagIndentAlignmentAndExplicitLines() {
        var page = new OneNotePage { PageSize = OneNotePageSize.IndexCard };
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 0.25, Y = 1, Width = 4.5 } };
        var paragraph = new OneNoteParagraph {
            List = new OneNoteListInfo { Ordered = false, Level = 1 }
        };
        paragraph.Style.Alignment = OneNoteParagraphAlignment.Right;
        paragraph.Tags.Add(new OneNoteTag { IsCheckable = true, IsCompleted = false });
        paragraph.AddMath(OfficeMath.Fraction(OfficeMath.Number("1"), OfficeMath.Number("2")));
        paragraph.Runs.Add(new OneNoteTextRun { Text = " value\nNext" });
        outline.Children.Add(paragraph);
        page.Outlines.Add(outline);

        OfficeDrawing drawing = page.ToDrawing();
        OfficeDrawingText[] text = drawing.Elements.OfType<OfficeDrawingText>().ToArray();
        string combined = string.Concat(text.Select(item => item.Text));
        OfficeDrawingText next = Assert.Single(text, item => item.Text == "Next");

        Assert.Contains("•", combined, StringComparison.Ordinal);
        Assert.Contains("☐", combined, StringComparison.Ordinal);
        Assert.True(next.X > 0.25D * OneNotePageRenderer.PointsPerHalfInch + 18D);
        Assert.True(next.Y > 1D * OneNotePageRenderer.PointsPerHalfInch);
    }

    [Fact]
    public void OversizedInlineMathRendersReadableFallbackInsideItsOutline() {
        const string expressionText = "MMMMMMMMMMMMMMMMMMMMMMMM";
        var page = new OneNotePage { PageSize = OneNotePageSize.IndexCard };
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 0.25D, Y = 0.5D, Width = 0.75D } };
        var paragraph = new OneNoteParagraph();
        paragraph.AddMath(OfficeMath.Identifier(expressionText));
        outline.Children.Add(paragraph);
        page.Outlines.Add(outline);

        OneNotePageVisualSnapshot snapshot = OneNotePageRenderer.CreateSnapshot(page);
        OfficeDrawingText fallback = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), item => item.Text == expressionText);
        double outlineRight = (0.25D + 0.75D) * OneNotePageRenderer.PointsPerHalfInch;

        Assert.True(fallback.X + fallback.Width <= outlineRight + 0.001D);
        Assert.Contains(snapshot.Diagnostics, diagnostic => diagnostic.Code == "ONENOTE_RENDER_MATH_CLIPPED");
    }

    [Fact]
    public void AutomaticCanvasIncludesPositionedOutlineChildren() {
        var page = new OneNotePage { PageSize = OneNotePageSize.Automatic, Width = 1D, Height = 1D };
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 0D, Y = 0D, Width = 2D } };
        var paragraph = new OneNoteParagraph { Layout = new OneNoteLayout { X = 8D, Y = 10D, Width = 2D } };
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Positioned child" });
        outline.Children.Add(paragraph);
        page.Outlines.Add(outline);
        var options = new OneNotePageRenderingOptions {
            AutomaticPageWidthPoints = 120D,
            AutomaticPageHeightPoints = 120D,
            AutomaticPagePaddingPoints = 24D,
            IncludeTitle = false
        };

        OfficeDrawing drawing = page.ToDrawing(options);
        OfficeDrawingRichText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingRichText>());

        Assert.True(text.X + text.Width + options.AutomaticPagePaddingPoints <= drawing.Width + 0.001D);
        Assert.True(text.Y + text.Height + options.AutomaticPagePaddingPoints <= drawing.Height + 0.001D);
    }

    [Fact]
    public void PositionedElementAtCanvasEdgeUsesTheActualRemainingWidth() {
        var page = new OneNotePage {
            PageSize = OneNotePageSize.Custom,
            Width = 4D,
            Height = 4D
        };
        var paragraph = new OneNoteParagraph {
            Layout = new OneNoteLayout { X = 3.99D, Y = 0.5D, Width = 1D }
        };
        paragraph.Runs.Add(new OneNoteTextRun { Text = "edge" });
        page.DirectContent.Add(paragraph);

        OfficeDrawing drawing = page.ToDrawing(new OneNotePageRenderingOptions { IncludeTitle = false });
        OfficeDrawingRichText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingRichText>());

        Assert.True(text.Width > 0D);
        Assert.True(text.X + text.Width <= drawing.Width + 0.001D);
    }

    [Fact]
    public void ParagraphChildrenRenderAfterRunsAndExplicitFlowHeightIsReserved() {
        var page = new OneNotePage { PageSize = OneNotePageSize.IndexCard };
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 0.25D, Y = 0.5D, Width = 4D } };
        var parent = new OneNoteParagraph { Layout = new OneNoteLayout { Height = 2D } };
        parent.Runs.Add(new OneNoteTextRun { Text = "Parent" });
        var nested = new OneNoteParagraph();
        nested.Runs.Add(new OneNoteTextRun { Text = "Nested" });
        parent.Children.Add(nested);
        var following = new OneNoteParagraph();
        following.Runs.Add(new OneNoteTextRun { Text = "Following" });
        outline.Children.Add(parent);
        outline.Children.Add(following);
        page.Outlines.Add(outline);

        OfficeDrawingRichText[] text = page.ToDrawing().Elements.OfType<OfficeDrawingRichText>().ToArray();
        OfficeDrawingRichText parentText = Assert.Single(text, item => item.Runs.Any(run => run.Text == "Parent"));
        OfficeDrawingRichText nestedText = Assert.Single(text, item => item.Runs.Any(run => run.Text == "Nested"));
        OfficeDrawingRichText followingText = Assert.Single(text, item => item.Runs.Any(run => run.Text == "Following"));

        Assert.True(nestedText.Y >= parentText.Y + parentText.Height);
        Assert.True(followingText.Y >= parentText.Y + 2D * OneNotePageRenderer.PointsPerHalfInch);
    }

    [Fact]
    public void PageAndElementRightToLeftDefaultsAlignBodyParagraphsToTheRight() {
        var page = new OneNotePage { PageSize = OneNotePageSize.IndexCard, RightToLeft = true };
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 0.25, Y = 1, Width = 4.5 } };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "مرحبا" });
        outline.Children.Add(paragraph);
        page.Outlines.Add(outline);

        OfficeDrawing drawing = page.ToDrawing();
        OfficeDrawingRichText body = Assert.Single(drawing.Elements.OfType<OfficeDrawingRichText>());

        Assert.Equal(OfficeTextAlignment.Right, body.Alignment);
    }

    [Fact]
    public void ParagraphSpacingCollapsesAndExactLineSpacingFlowsToTheNextParagraph() {
        var page = new OneNotePage { PageSize = OneNotePageSize.IndexCard };
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 0.25, Y = 0.5, Width = 4.5 } };
        var first = new OneNoteParagraph();
        first.Style.ExactLineSpacing = 0.75D;
        first.Style.SpaceAfter = 0.5D;
        first.Runs.Add(new OneNoteTextRun { Text = "First line\nSecond line" });
        var second = new OneNoteParagraph();
        second.Style.SpaceBefore = 1D;
        second.Runs.Add(new OneNoteTextRun { Text = "Next paragraph" });
        outline.Children.Add(first);
        outline.Children.Add(second);
        page.Outlines.Add(outline);

        OfficeDrawingRichText[] text = page.ToDrawing().Elements.OfType<OfficeDrawingRichText>().ToArray();

        Assert.Equal(2, text.Length);
        Assert.Equal(27D, text[0].LineHeight!.Value, 6);
        Assert.Equal(36D, text[1].Y - (text[0].Y + text[0].Height), 6);
    }

    [Fact]
    public void RichTextHeightWrapsEachHardLineIndependently() {
        var page = new OneNotePage { PageSize = OneNotePageSize.IndexCard };
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 0.25, Y = 0.5, Width = 4 } };
        var paragraph = new OneNoteParagraph();
        string wideLine = new string('M', 30);
        paragraph.Runs.Add(new OneNoteTextRun { Text = wideLine + "\n" + wideLine });
        outline.Children.Add(paragraph);
        page.Outlines.Add(outline);

        OfficeDrawingRichText text = Assert.Single(page.ToDrawing().Elements.OfType<OfficeDrawingRichText>());

        Assert.True(text.Height >= text.LineHeight!.Value * 4D);
    }

    [Fact]
    public void NarrowTableRowsMeasureWrappedCellContentAndAlwaysUseHalfInchColumnWidths() {
        var page = new OneNotePage { PageSize = OneNotePageSize.Letter };
        var table = new OneNoteTable { Layout = new OneNoteLayout { X = 0.5, Y = 0.5, Width = 5 }, BordersVisible = true };
        table.ColumnWidths.Add(50D);
        table.ColumnWidths.Add(50D);
        var firstRow = new OneNoteTableRow();
        firstRow.Cells.Add(CellWithText(string.Join(" ", Enumerable.Repeat("wrapped", 20))));
        firstRow.Cells.Add(CellWithText("side"));
        var secondRow = new OneNoteTableRow();
        secondRow.Cells.Add(CellWithText("next"));
        secondRow.Cells.Add(CellWithText("row"));
        table.Rows.Add(firstRow);
        table.Rows.Add(secondRow);
        page.DirectContent.Add(table);

        OfficeDrawingShape[] frames = page.ToDrawing().Elements.OfType<OfficeDrawingShape>()
            .Where(shape => shape.Shape.StrokeWidth == 0.75D)
            .ToArray();

        Assert.Equal(4, frames.Length);
        Assert.True(frames[0].Shape.Height > 32D);
        Assert.Equal(90D, frames[1].X - frames[0].X, 6);
        Assert.True(frames[2].Y >= frames[0].Y + frames[0].Shape.Height);
    }

    [Fact]
    public void AutomaticCanvasExpandsForSoftWrappedParagraphsAndMeasuredTables() {
        var page = new OneNotePage { PageSize = OneNotePageSize.Automatic };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = string.Join(" ", Enumerable.Repeat("automatic wrapping", 600)) });
        page.DirectContent.Add(paragraph);

        OfficeDrawing drawing = page.ToDrawing(new OneNotePageRenderingOptions {
            AutomaticPageWidthPoints = 240D,
            AutomaticPageHeightPoints = 120D,
            AutomaticPagePaddingPoints = 24D,
            IncludeTitle = false
        });
        OfficeDrawingRichText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingRichText>());

        Assert.True(drawing.Height > 792D);
        Assert.True(text.Y + text.Height + 24D <= drawing.Height + 0.001D);
    }

    private static OneNoteTableCell CellWithText(string text) {
        var cell = new OneNoteTableCell();
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = text });
        cell.Content.Add(paragraph);
        return cell;
    }

    private static OneNotePage CreateVisualPage(string title = "Premium canvas") {
        var page = new OneNotePage { Title = title, PageSize = OneNotePageSize.IndexCard };
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 0.35, Y = 1.25, Width = 5.3 } };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Rich " });
        var bold = new OneNoteTextRun { Text = "OneNote" };
        bold.Style.Bold = true;
        bold.Style.ColorArgb = 0xFF2255AA;
        paragraph.Runs.Add(bold);
        paragraph.Runs.Add(new OneNoteTextRun { Text = " canvas with " });
        paragraph.AddMath(OfficeMath.Fraction(OfficeMath.Number("1"), OfficeMath.Number("2")));
        outline.Children.Add(paragraph);

        var image = new OneNoteImage {
            FileName = "blue.png",
            MediaType = "image/png",
            AltText = "Blue sample",
            WidthHalfInches = 1.2,
            HeightHalfInches = 0.6,
            Payload = OneNoteBinaryPayload.FromBytes(OfficePngWriter.Encode(new OfficeRasterImage(8, 4, OfficeColor.CornflowerBlue)))
        };
        outline.Children.Add(image);

        var table = new OneNoteTable { BordersVisible = true };
        table.ColumnWidths.Add(2.4);
        table.ColumnWidths.Add(2.4);
        var row = new OneNoteTableRow();
        var left = new OneNoteTableCell { ShadingColorArgb = 0xFFF1F5FA };
        var leftParagraph = new OneNoteParagraph();
        leftParagraph.Runs.Add(new OneNoteTextRun { Text = "Ink" });
        left.Content.Add(leftParagraph);
        var right = new OneNoteTableCell();
        var rightParagraph = new OneNoteParagraph();
        rightParagraph.Runs.Add(new OneNoteTextRun { Text = "Math" });
        right.Content.Add(rightParagraph);
        row.Cells.Add(left);
        row.Cells.Add(right);
        table.Rows.Add(row);
        outline.Children.Add(table);

        var ink = new OneNoteInk();
        var stroke = new OfficeInkStroke {
            Color = OfficeColor.Crimson,
            Width = 0.04,
            Height = 0.04,
            FitToCurve = true
        };
        stroke.AddPoint(0.2, 0.1, 0.4).AddPoint(1.2, 0.45, 1D).AddPoint(2.2, 0.15, 0.6);
        ink.Ink.Add(stroke);
        outline.Children.Add(ink);
        outline.Children.Add(new OneNoteEmbeddedFile { FileName = "brief.pdf", MediaType = "application/pdf", Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3 }) });
        page.Outlines.Add(outline);
        return page;
    }
}
