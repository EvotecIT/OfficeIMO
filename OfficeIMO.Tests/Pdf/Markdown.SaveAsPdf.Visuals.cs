using OfficeIMO.Drawing;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class MarkdownSaveAsPdfVisualTests {
    [Fact]
    public void ToPdfDocument_MarkdownImageBlock_RendersDataUriImageAsStyledFigure() {
        string dataUri = CreateDataUriPng();
        MarkdownDoc document = MarkdownDoc
            .Create()
            .H1("Visual report")
            .Image(dataUri, "Operational badge", "Badge", width: 42, height: 28)
            .Caption("Figure 1. Operational badge");

        byte[] bytes = document.ToPdfDocument(CreateVisualOptions()).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains("Figure 1. Operational badge", text, StringComparison.Ordinal);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(bytes), image => image.IsImageFile && image.MimeType == "image/png");
    }

    [Fact]
    public void ToPdfDocument_MarkdownImageOnlyParagraph_RendersImageInsteadOfTextPlaceholder() {
        string dataUri = CreateDataUriPng();
        MarkdownDoc document = MarkdownDoc
            .Create()
            .Add(new ParagraphBlock(new InlineSequence().Image("Inline badge", dataUri, "Inline badge caption")));

        byte[] bytes = document.ToPdfDocument(CreateVisualOptions()).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(bytes), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.Contains("Inline badge caption", text, StringComparison.Ordinal);
        Assert.DoesNotContain("[Image:", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownImageOnlyParagraphInsideQuote_RendersOutsidePanelWithoutThrowing() {
        string dataUri = CreateDataUriPng();
        var quote = new QuoteBlock();
        quote.Children.Add(new ParagraphBlock(new InlineSequence().Image("Inline badge", dataUri, "Inline badge caption")));
        MarkdownDoc document = MarkdownDoc
            .Create()
            .Add(quote);

        byte[] bytes = document.ToPdfDocument(CreateVisualOptions()).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(bytes), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.Contains("Inline badge caption", text, StringComparison.Ordinal);
        Assert.DoesNotContain("[Image:", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownImageOnlyListContinuationInsideQuote_RendersOutsidePanelWithoutThrowing() {
        string dataUri = CreateDataUriPng();
        var item = ListItem.Text("Evidence");
        item.AdditionalParagraphs.Add(new InlineSequence().Image("Inline badge", dataUri, "Inline badge caption"));
        var list = new UnorderedListBlock();
        list.Items.Add(item);
        var quote = new QuoteBlock();
        quote.Children.Add(list);
        MarkdownDoc document = MarkdownDoc
            .Create()
            .Add(quote);

        byte[] bytes = document.ToPdfDocument(CreateVisualOptions()).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(bytes), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.Contains("Inline badge caption", text, StringComparison.Ordinal);
        Assert.DoesNotContain("[Image:", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownImageOnlyParagraph_WarnsAndFallsBackWhenBytesAreUnsupportedByPdfRenderer() {
        const string imageUrl = "https://example.test/badge.gif";
        MarkdownDoc document = MarkdownDoc
            .Create()
            .Add(new ParagraphBlock(new InlineSequence().Image("Remote badge", imageUrl, "Remote badge caption")));
        var options = CreateVisualOptions();
        options.RemoteImageResolver = _ => CreateGifBytes();

        byte[] bytes = document.ToPdfDocument(options).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains(options.Warnings, warning => warning.Code == "UnsupportedImage" && warning.Source == imageUrl);
        Assert.Contains("[Image unavailable: Remote badge]", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownLinkedImageWithBlankAlt_RendersWithoutBlankLinkContentsException() {
        string dataUri = CreateDataUriPng();
        string markdown = "[![](" + dataUri + ")](https://example.test/report)\n";

        byte[] bytes = markdown.ToPdfDocument(CreateVisualOptions()).ToBytes();

        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(bytes), image => image.IsImageFile && image.MimeType == "image/png");
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_RendersChartVisualInsteadOfJsonCodePanel() {
        const string markdown = """
# Quarterly report

```chart
{
  "type": "bar",
  "title": "Quarter revenue",
  "data": {
    "labels": ["Q1", "Q2", "Q3"],
    "datasets": [
      { "label": "Actual", "data": [10, 14, 19], "backgroundColor": "#2563EB" }
    ]
  },
  "width": 360,
  "height": 220
}
```
_Figure 2. Revenue chart_
""";

        var options = CreateVisualOptions();
        byte[] bytes = markdown.ToPdfDocument(options).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "UnsupportedChartFence");
        Assert.Contains("Quarter revenue", text, StringComparison.Ordinal);
        Assert.Contains("Actual", text, StringComparison.Ordinal);
        Assert.Contains("Q1", text, StringComparison.Ordinal);
        Assert.Contains("Figure 2. Revenue chart", text, StringComparison.Ordinal);
        Assert.DoesNotContain("\"datasets\"", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFenceInsideCallout_RendersOutsidePanelWithoutThrowing() {
        var callout = new CalloutBlock("note", "Visual", new IMarkdownBlock[] {
            new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "title": "Quarter revenue",
  "data": {
    "labels": ["Q1", "Q2"],
    "datasets": [
      { "label": "Actual", "data": [10, 14] }
    ]
  }
}
""")
        });
        MarkdownDoc document = MarkdownDoc
            .Create()
            .Add(callout);

        var options = CreateVisualOptions();
        byte[] bytes = document.ToPdfDocument(options).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "UnsupportedChartFence");
        Assert.Contains("Visual", text, StringComparison.Ordinal);
        Assert.Contains("Quarter revenue", text, StringComparison.Ordinal);
        Assert.Contains("Actual", text, StringComparison.Ordinal);
        Assert.DoesNotContain("\"datasets\"", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_UsesFenceTitleWhenJsonTitleIsMissing() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart title=\"Fence revenue\"", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Equal("Fence revenue", snapshot!.Title);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_MapsChartJsHorizontalBar() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": { "indexAxis": "y" },
  "data": {
    "labels": ["Backlog", "Done"],
    "datasets": [
      { "label": "Items", "data": [3, 7] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Equal(OfficeChartKind.BarClustered, snapshot!.ChartKind);
    }

    [Fact]
    public void ToPdfDocument_MarkdownScatterChartFence_PreservesExplicitXValuesAndMissingYValues() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "scatter",
  "data": {
    "datasets": [
      { "label": "Samples", "data": [
        { "x": 10, "y": 2 },
        { "x": 25, "y": null },
        { "x": 40, "y": 8 }
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeChartSeries series = Assert.Single(snapshot!.Data.Series);
        Assert.Equal(new[] { 10D, 25D, 40D }, series.XValues);
        Assert.Equal(2D, series.Values[0]);
        Assert.True(double.IsNaN(series.Values[1]));
        Assert.Equal(8D, series.Values[2]);
    }

    [Fact]
    public void ToPdfDocument_MarkdownScatterChartFence_PreservesPointCountWhenLabelsAreMetadata() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "scatter",
  "data": {
    "labels": ["Reference"],
    "datasets": [
      { "label": "Samples", "data": [
        { "x": 10, "y": 2 },
        { "x": 25, "y": 4 }
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeChartSeries series = Assert.Single(snapshot!.Data.Series);
        Assert.Equal(new[] { 10D, 25D }, series.XValues);
        Assert.Equal(new[] { 2D, 4D }, series.Values);
        Assert.Equal(2, snapshot.Data.Categories.Count);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_PreservesObjectPointCategoriesWhenLabelsAreMissing() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "datasets": [
      { "label": "Actual", "data": [
        { "x": "Q1", "y": 10 },
        { "x": "Q2", "y": 14 }
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Equal(new[] { "Q1", "Q2" }, snapshot!.Data.Categories);
        Assert.Equal(new[] { 10D, 14D }, Assert.Single(snapshot.Data.Series).Values);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_MergesObjectPointCategoriesAcrossDatasetsWhenLabelsAreMissing() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "datasets": [
      { "label": "Actual", "data": [
        { "x": "Q1", "y": 10 }
      ] },
      { "label": "Forecast", "data": [
        { "x": "Q2", "y": 14 }
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Equal(new[] { "Q1", "Q2" }, snapshot!.Data.Categories);
        Assert.Equal(10D, snapshot.Data.Series[0].Values[0]);
        Assert.True(double.IsNaN(snapshot.Data.Series[0].Values[1]));
        Assert.True(double.IsNaN(snapshot.Data.Series[1].Values[0]));
        Assert.Equal(14D, snapshot.Data.Series[1].Values[1]);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_AlignsObjectPointCategoriesToExplicitLabels() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1", "Q2"],
    "datasets": [
      { "label": "Actual", "data": [
        { "x": "Q2", "y": 14 }
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeChartSeries series = Assert.Single(snapshot!.Data.Series);
        Assert.True(double.IsNaN(series.Values[0]));
        Assert.Equal(14D, series.Values[1]);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_PreservesMixedScalarAndObjectPointPositionsWithExplicitLabels() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1", "Q2"],
    "datasets": [
      { "label": "Actual", "data": [
        10,
        { "x": "Q2", "y": 14 }
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeChartSeries series = Assert.Single(snapshot!.Data.Series);
        Assert.Equal(10D, series.Values[0]);
        Assert.Equal(14D, series.Values[1]);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_UsesMaximumScalarDatasetLengthWhenLabelsAreMissing() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "datasets": [
      { "label": "Short", "data": [1] },
      { "label": "Long", "data": [2, 3] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Equal(new[] { "1", "2" }, snapshot!.Data.Categories);
        Assert.Equal(1D, snapshot.Data.Series[0].Values[0]);
        Assert.True(double.IsNaN(snapshot.Data.Series[0].Values[1]));
        Assert.Equal(new[] { 2D, 3D }, snapshot.Data.Series[1].Values);
    }

    [Fact]
    public void ToPdfDocument_MarkdownScatterChartFence_ReadsTuplePointArrays() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "scatter",
  "data": {
    "datasets": [
      { "label": "Samples", "data": [[10, 2], [25, 4]] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeChartSeries series = Assert.Single(snapshot!.Data.Series);
        Assert.Equal(new[] { 10D, 25D }, series.XValues);
        Assert.Equal(new[] { 2D, 4D }, series.Values);
    }

    [Fact]
    public void ToPdfDocument_MarkdownScatterChartFence_PreservesRootTupleValueXCoordinates() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "scatter",
  "values": [[10, 2], [25, 4]]
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeChartSeries series = Assert.Single(snapshot!.Data.Series);
        Assert.Equal(new[] { 10D, 25D }, series.XValues);
        Assert.Equal(new[] { 2D, 4D }, series.Values);
    }

    [Fact]
    public void ToPdfDocument_MarkdownScatterChartFence_WarnsWhenExplicitXValuesCannotRenderPoint() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "scatter",
  "data": {
    "datasets": [
      { "label": "Samples", "data": [
        { "x": "2026-01-01", "y": 4 },
        { "x": "2026-01-02", "y": 8 }
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("finite X/Y point", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownScatterChartFence_IgnoresInvalidXPointValuesBeforeRangeCalculation() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "scatter",
  "data": {
    "datasets": [
      { "label": "Samples", "data": [
        { "x": "bad", "y": 1000000 },
        { "x": 1, "y": 2 }
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeChartSeries series = Assert.Single(snapshot!.Data.Series);
        Assert.True(double.IsNaN(series.XValues![0]));
        Assert.True(double.IsNaN(series.Values[0]));
        Assert.Equal(1D, series.XValues![1]);
        Assert.Equal(2D, series.Values[1]);
    }

    [Fact]
    public void ToPdfDocument_MarkdownScatterChartFence_DoesNotConnectPointsUnlessShowLineIsEnabled() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "scatter",
  "data": {
    "datasets": [
      { "label": "Samples", "data": [
        { "x": 1, "y": 2 },
        { "x": 2, "y": 4 }
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.False(snapshot!.Layout.ConnectScatterPoints);
        Assert.False(Assert.Single(snapshot.Data.Series).ConnectLine);
    }

    [Fact]
    public void ToPdfDocument_MarkdownScatterChartFence_ConnectsPointsWhenShowLineIsEnabled() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "scatter",
  "showLine": true,
  "data": {
    "datasets": [
      { "label": "Samples", "data": [
        { "x": 1, "y": 2 },
        { "x": 2, "y": 4 }
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.True(snapshot!.Layout.ConnectScatterPoints);
        Assert.True(Assert.Single(snapshot.Data.Series).ConnectLine);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_PreservesMissingArrayValuesAsNaN() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "data": {
    "labels": ["A", "B", "C", "D"],
    "datasets": [
      { "label": "Actual", "data": [1, null, 3] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeChartSeries series = Assert.Single(snapshot!.Data.Series);
        Assert.Equal(1D, series.Values[0]);
        Assert.True(double.IsNaN(series.Values[1]));
        Assert.Equal(3D, series.Values[2]);
        Assert.True(double.IsNaN(series.Values[3]));
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsWhenSeriesContainsOnlyMissingValues() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [null] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("finite", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsWhenPageIsTooNarrowForNativeChartRenderer() {
        var options = CreateVisualOptions();
        options.PdfOptions!.PageWidth = 220;
        options.PdfOptions.MarginLeft = 36;
        options.PdfOptions.MarginRight = 36;
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, options, out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("240", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsWhenPageIsTooShortForNativeChartRenderer() {
        var options = CreateVisualOptions();
        options.PdfOptions!.PageHeight = 180;
        options.PdfOptions.MarginTop = 36;
        options.PdfOptions.MarginBottom = 36;
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, options, out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("150", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsWhenFigureSpacingLeavesPageTooShortForNativeChartRenderer() {
        var options = CreateVisualOptions();
        options.PdfOptions!.PageHeight = 226;
        options.PdfOptions.MarginTop = 36;
        options.PdfOptions.MarginBottom = 36;
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, options, out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("figure spacing", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownRadarChartFence_WarnsWhenCategoryCountCannotRenderRadar() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "radar",
  "data": {
    "labels": ["Quality", "Speed"],
    "datasets": [
      { "label": "Score", "data": [8, 9] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("three categories", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownRadarChartFence_WarnsWhenFiniteValuesCannotDrawSegment() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "radar",
  "data": {
    "labels": ["Quality", "Speed", "Cost"],
    "datasets": [
      { "label": "Score", "data": [8, null, null] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("drawable", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownAreaChartFence_WarnsWhenCategoryCountCannotRenderArea() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "area",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("two categories", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownAreaChartFence_WarnsWhenFiniteValuesCannotDrawRun() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "area",
  "data": {
    "labels": ["Q1", "Q2"],
    "datasets": [
      { "label": "Actual", "data": [10, null] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("adjacent finite", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownAreaChartFence_ExcludesSeriesWithoutDrawableRun() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "area",
  "data": {
    "labels": ["Q1", "Q2"],
    "datasets": [
      { "label": "Visible", "data": [1, 2] },
      { "label": "Invisible", "data": [1000, null] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeChartSeries series = Assert.Single(snapshot!.Data.Series);
        Assert.Equal("Visible", series.Name);
        Assert.Equal(new[] { 1D, 2D }, series.Values);
    }

    [Fact]
    public void ToPdfDocument_MarkdownDoughnutChartFence_PreservesMultipleDatasets() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "doughnut",
  "data": {
    "labels": ["Passed", "Failed"],
    "datasets": [
      { "label": "Outer", "data": [8, 2] },
      { "label": "Inner", "data": [6, 4] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Equal(OfficeChartKind.Doughnut, snapshot!.ChartKind);
        Assert.Equal(2, snapshot.Data.Series.Count);
        Assert.Equal("Outer", snapshot.Data.Series[0].Name);
        Assert.Equal("Inner", snapshot.Data.Series[1].Name);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_SkipsHiddenChartJsDatasets() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Hidden", "hidden": true, "data": [1000] },
      { "label": "Visible", "data": [2] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeChartSeries series = Assert.Single(snapshot!.Data.Series);
        Assert.Equal("Visible", series.Name);
        Assert.Equal(2D, Assert.Single(series.Values));
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_RespectsDisabledChartJsLegend() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "plugins": {
      "legend": { "display": false }
    }
  },
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.False(snapshot!.Layout.ShowLegend);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_RespectsDisabledChartJsTitlePlugin() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "plugins": {
      "title": { "display": false, "text": "Hidden title" }
    }
  },
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Null(snapshot!.Title);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_IgnoresChartJsTitlePluginUnlessDisplayIsEnabled() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "plugins": {
      "title": { "text": "Metadata title" }
    }
  },
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Null(snapshot!.Title);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_MapsChartJsStackedBars() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "scales": {
      "x": { "stacked": true },
      "y": { "stacked": true }
    }
  },
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] },
      { "label": "Forecast", "data": [12] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Equal(OfficeChartKind.ColumnStacked, snapshot!.ChartKind);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_MapsChartJsStackedLines() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "options": {
    "scales": {
      "y": { "stacked": true }
    }
  },
  "data": {
    "labels": ["Q1", "Q2"],
    "datasets": [
      { "label": "Actual", "data": [10, 12] },
      { "label": "Forecast", "data": [8, 11] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Equal(OfficeChartKind.LineStacked, snapshot!.ChartKind);
    }

    [Fact]
    public void ToPdfDocument_MarkdownPieChartFence_WarnsWhenNoPositiveFiniteSliceCanRender() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "pie",
  "data": {
    "labels": ["Passed", "Failed", "Skipped"],
    "datasets": [
      { "label": "Status", "data": [0, -2, null] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("positive finite slice", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_CustomReaderOptionsWithoutSemanticChartExtension_LeavesChartFenceAsCodePanel() {
        const string markdown = """
```chart
{
  "type": "bar",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  }
}
```
""";

        var options = CreateVisualOptions();
        options.ReaderOptions = MarkdownReaderOptions.CreateOfficeIMOProfile();

        byte[] bytes = markdown.ToPdfDocument(options).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "UnsupportedChartFence");
        Assert.Contains("\"datasets\"", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WithDecorativeFigureDrawingStyle_RendersWithChartAltText() {
        const string markdown = """
```chart
{
  "type": "bar",
  "title": "Decorative source style",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  }
}
```
""";

        MarkdownPdfVisualTheme theme = MarkdownPdfVisualTheme.Report();
        MarkdownPdfFigureStyle figureStyle = theme.FigureStyle!;
        figureStyle.DrawingStyle = new PdfCore.PdfDrawingStyle {
            Align = PdfCore.PdfAlign.Center,
            Decorative = true,
            KeepWithNext = true,
            SpacingBefore = 2,
            SpacingAfter = 4
        };
        theme.FigureStyle = figureStyle;
        var options = CreateVisualOptions();
        options.VisualTheme = theme;

        byte[] bytes = markdown.ToPdfDocument(options).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "UnsupportedChartFence");
        Assert.Contains("Decorative source style", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_InvalidMarkdownChartFence_WarnsAndFallsBackToSemanticPanel() {
        const string markdown = """
```chart
{ invalid json
```
""";

        var options = CreateVisualOptions();
        byte[] bytes = markdown.ToPdfDocument(options).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains(options.Warnings, warning => warning.Code == "UnsupportedChartFence");
        Assert.Contains("{ invalid json", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_UnsupportedVisualFence_UsesSemanticFallbackPanelWithWarning() {
        const string markdown = """
```mermaid
graph TD
A[Markdown AST] --> B[OfficeIMO PDF]
```
_Figure 3. Flow fallback_
""";

        var options = CreateVisualOptions();
        byte[] bytes = markdown.ToPdfDocument(options).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains(options.Warnings, warning => warning.Code == "UnsupportedSemanticFence" && warning.Source == MarkdownSemanticKinds.Mermaid);
        Assert.Contains("mermaid", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("graphTD", text, StringComparison.Ordinal);
        Assert.Contains("Figure 3. Flow fallback", text, StringComparison.Ordinal);
    }

    private static MarkdownPdfSaveOptions CreateVisualOptions() => new MarkdownPdfSaveOptions {
        VisualTheme = MarkdownPdfVisualTheme.Report(),
        PdfOptions = new PdfCore.PdfOptions {
            CompressContentStreams = false,
            PageWidth = 420,
            PageHeight = 420,
            MarginLeft = 36,
            MarginRight = 36,
            MarginTop = 36,
            MarginBottom = 36
        }
    };

    private static string CreateDataUriPng() {
        string base64 = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 1));
        return "data:image/png;base64," + base64;
    }

    private static byte[] CreateGifBytes() => Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");
}
