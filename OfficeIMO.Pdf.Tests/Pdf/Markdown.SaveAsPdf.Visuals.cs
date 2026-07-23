using OfficeIMO.Drawing;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System;
using System.Reflection;
using System.Text;
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
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

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
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(bytes), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.DoesNotContain("Inline badge caption", text, StringComparison.Ordinal);
        Assert.DoesNotContain("[Image:", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownImageOnlyParagraphInsideQuote_RendersOutsidePanelWithoutThrowing() {
        string dataUri = CreateDataUriPng();
        var quote = new QuoteBlock();
        quote.ChildBlocks.Add(new ParagraphBlock(new InlineSequence().Image("Inline badge", dataUri, "Inline badge caption")));
        MarkdownDoc document = MarkdownDoc
            .Create()
            .Add(quote);

        byte[] bytes = document.ToPdfDocument(CreateVisualOptions()).ToBytes();
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(bytes), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.DoesNotContain("Inline badge caption", text, StringComparison.Ordinal);
        Assert.DoesNotContain("[Image:", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Quote", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownImageOnlyListContinuationInsideQuote_RendersOutsidePanelWithoutThrowing() {
        string dataUri = CreateDataUriPng();
        var item = ListItem.Text("Evidence");
        item.AdditionalParagraphs.Add(new InlineSequence().Image("Inline badge", dataUri, "Inline badge caption"));
        var list = new UnorderedListBlock();
        list.Items.Add(item);
        var quote = new QuoteBlock();
        quote.ChildBlocks.Add(list);
        MarkdownDoc document = MarkdownDoc
            .Create()
            .Add(quote);

        byte[] bytes = document.ToPdfDocument(CreateVisualOptions()).ToBytes();
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(bytes), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.DoesNotContain("Inline badge caption", text, StringComparison.Ordinal);
        Assert.DoesNotContain("[Image:", text, StringComparison.Ordinal);
    }

    [Fact]
    public void FromMarkdownTheme_AppliesSharedHeadingAndTextColorsToPdfDocumentTheme() {
        MarkdownVisualTheme sharedTheme = MarkdownVisualTheme.Report()
            .WithColors(heading: "#123456", text: "#654321");

        MarkdownPdfStyle pdfTheme = MarkdownPdfStyle.FromMarkdownTheme(sharedTheme);
        PdfCore.PdfTheme documentTheme = pdfTheme.DocumentTheme!;

        Assert.Equal(PdfCore.PdfColor.FromRgb(0x65, 0x43, 0x21), documentTheme.TextStyle!.Color);
        Assert.Equal(PdfCore.PdfColor.FromRgb(0x12, 0x34, 0x56), documentTheme.HeadingStyles!.Level1!.Color);
        Assert.Equal(PdfCore.PdfColor.FromRgb(0x12, 0x34, 0x56), documentTheme.HeadingStyles.Level2!.Color);
        Assert.Equal(PdfCore.PdfColor.FromRgb(0x12, 0x34, 0x56), documentTheme.HeadingStyles.Level3!.Color);
    }

    [Fact]
    public void FromMarkdownTheme_AppliesSharedBackgroundToPdfPageDecoration() {
        MarkdownVisualTheme sharedTheme = MarkdownVisualTheme.Report()
            .WithColors(background: "#112233");

        MarkdownPdfStyle pdfTheme = MarkdownPdfStyle.FromMarkdownTheme(sharedTheme);

        Assert.Equal(PdfCore.PdfColor.FromRgb(0x11, 0x22, 0x33), pdfTheme.PageDecoration!.BackgroundColor);
    }

    [Fact]
    public void FromMarkdownTheme_TreatsTransparentSharedBackgroundAsNoPageDecoration() {
        MarkdownVisualTheme sharedTheme = MarkdownVisualTheme.Report()
            .WithColors(background: "Transparent");

        MarkdownPdfStyle pdfTheme = MarkdownPdfStyle.FromMarkdownTheme(sharedTheme);

        Assert.Null(pdfTheme.PageDecoration);
    }

    [Fact]
    public void ToPdfDocument_SharedMarkdownTheme_DoesNotOverrideExplicitPdfBackground() {
        MarkdownVisualTheme sharedTheme = MarkdownVisualTheme.Report()
            .WithColors(background: "#112233");
        var options = new MarkdownPdfSaveOptions {
            Theme = sharedTheme,
            PdfOptions = new PdfCore.PdfOptions {
                BackgroundColor = PdfCore.PdfColor.FromRgb(0xaa, 0xbb, 0xcc)
            }
        };

        PdfCore.PdfDocument pdf = OfficeIMO.Markdown.MarkdownReader.Parse("# Themed").ToPdfDocument(options);

        Assert.Equal(PdfCore.PdfColor.FromRgb(0xaa, 0xbb, 0xcc), GetPdfOptions(pdf).BackgroundColor);
    }

    [Fact]
    public void ToPdfDocument_SharedMarkdownTheme_AppliesHeadingColorToLowerLevelHeadings() {
        MarkdownVisualTheme sharedTheme = MarkdownVisualTheme.Report()
            .WithColors(heading: "#ff0000", text: "#000000");
        var options = CreateVisualOptions();
        options.Style = null;
        options.Theme = sharedTheme;

        byte[] bytes = OfficeIMO.Markdown.MarkdownReader.Parse("#### Lower heading").ToPdfDocument(options).ToBytes();
        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("Lower heading", text, StringComparison.Ordinal);
        Assert.Contains("1 0 0 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void FromMarkdownTheme_HonorsDisabledTableHeaderEmphasis() {
        MarkdownVisualTheme sharedTheme = MarkdownVisualTheme.Report()
            .WithColors(tableHeaderBackground: "#fedcba", tableHeaderText: "#010203")
            .WithTable(table => table.EmphasizeHeader = false);

        MarkdownPdfStyle pdfTheme = MarkdownPdfStyle.FromMarkdownTheme(sharedTheme);
        PdfCore.PdfTableStyle style = pdfTheme.TableStyle!;

        Assert.Null(style.HeaderFill);
        Assert.Null(style.HeaderTextColor);
    }

    [Fact]
    public void FromMarkdownTheme_TreatsTransparentSharedFillsAsAbsent() {
        MarkdownVisualTheme sharedTheme = MarkdownVisualTheme.Report()
            .WithColors(codeBackground: "#00000000", tableStripeBackground: "Transparent")
            .WithTable(table => table.UseRowStripes = true);

        MarkdownPdfStyle pdfTheme = MarkdownPdfStyle.FromMarkdownTheme(sharedTheme);

        Assert.Null(pdfTheme.TableStyle!.RowStripeFill);
        Assert.Null(pdfTheme.CodeBlockPanelStyle!.Background);
    }

    [Fact]
    public void ToPdfDocument_QuoteWithImageOnlyChildKeepsTextInsideQuotePanel() {
        string markdown = $"""
> Intro text.
>
> ![Badge]({CreateDataUriPng()})
>
> Outro text.
""";
        MarkdownPdfStyle visualTheme = MarkdownPdfStyle.Plain();
        visualTheme.QuotePanelStyle = new PdfCore.PanelStyle {
            Background = PdfCore.PdfColor.FromRgb(0xff, 0, 0),
            PaddingX = 8,
            PaddingY = 6
        };
        var options = CreateVisualOptions();
        options.Style = visualTheme;

        byte[] bytes = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdfDocument(options).ToBytes();
        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("Intro text.", text, StringComparison.Ordinal);
        Assert.Contains("Outro text.", text, StringComparison.Ordinal);
        Assert.Contains("1 0 0 rg", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_FrontMatterPdfThemeOverridesGenericTheme() {
        MarkdownDoc document = MarkdownDoc.Create()
            .FrontMatter(new { theme = "report", pdfTheme = "technicalDocument" })
            .H1("Themed document");
        var options = new MarkdownPdfSaveOptions {
            UseFrontMatterTheme = true
        };
        MethodInfo method = typeof(MarkdownPdfConverterExtensions).GetMethod("ResolveVisualTheme", BindingFlags.NonPublic | BindingFlags.Static)!;

        MarkdownPdfStyle resolved = (MarkdownPdfStyle)method.Invoke(null, new object[] { document, options })!;

        Assert.Equal("TechnicalDocument", resolved.Name);
    }

    [Fact]
    public void ToPdfDocument_MarkdownGifImage_NormalizesThroughSharedRasterEngine() {
        const string imageUrl = "https://example.test/badge.gif";
        MarkdownDoc document = MarkdownDoc
            .Create()
            .Add(new ParagraphBlock(new InlineSequence().Image("Remote badge", imageUrl, "Remote badge caption")));
        var options = CreateVisualOptions();
        options.ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost();
        options.RemoteImageResolver = _ => CreateGifBytes();

        PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
        byte[] bytes = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "UnsupportedImage" && warning.Source == imageUrl);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(bytes), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.DoesNotContain("[Image unavailable: Remote badge]", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownLinkedImageWithBlankAlt_RendersWithoutBlankLinkContentsException() {
        string dataUri = CreateDataUriPng();
        string markdown = "[![](" + dataUri + ")](https://example.test/report)\n";

        byte[] bytes = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdfDocument(CreateVisualOptions()).ToBytes();

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
        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, MarkdownPdfSemanticBlocks.CreateReaderOptions()).ToPdfDocumentResult(options);
        byte[] bytes = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "UnsupportedChartFence");
        Assert.Contains("Quarter revenue", text, StringComparison.Ordinal);
        Assert.Contains("Actual", text, StringComparison.Ordinal);
        Assert.Contains("Q1", text, StringComparison.Ordinal);
        Assert.Contains("Figure 2. Revenue chart", text, StringComparison.Ordinal);
        Assert.DoesNotContain("\"datasets\"", text, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownChartJson_RejectsExcessiveDepthAndValueCounts() {
        string deeplyNested = new string('[', MarkdownPdfJsonValue.MaximumNestingDepth + 2) +
            "0" +
            new string(']', MarkdownPdfJsonValue.MaximumNestingDepth + 2);
        Assert.Throws<FormatException>(() => MarkdownPdfJsonValue.Parse(deeplyNested));

        string tooManyValues = "[" + string.Join(",", Enumerable.Repeat("0", MarkdownPdfJsonValue.MaximumValueNodes + 1)) + "]";
        Assert.Throws<FormatException>(() => MarkdownPdfJsonValue.Parse(tooManyValues));

        string tooLarge = new string(' ', MarkdownPdfJsonValue.MaximumInputCharacters + 1);
        Assert.Throws<FormatException>(() => MarkdownPdfJsonValue.Parse(tooLarge));
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
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

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
    public void ToPdfDocument_MarkdownChartFence_ReadsHorizontalBarObjectValuesFromX() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": { "indexAxis": "y" },
  "data": {
    "datasets": [
      { "label": "Items", "data": [
        { "x": 10, "y": "Backlog" },
        { "x": 14, "y": "Done" }
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Equal(new[] { "Backlog", "Done" }, snapshot!.Data.Categories);
        Assert.Equal(new[] { 10D, 14D }, Assert.Single(snapshot.Data.Series).Values);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_ReadsHorizontalBarTupleValuesFromX() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": { "indexAxis": "y" },
  "data": {
    "datasets": [
      { "label": "Items", "data": [
        [10, "Backlog"],
        [14, "Done"]
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Equal(new[] { "Backlog", "Done" }, snapshot!.Data.Categories);
        Assert.Equal(new[] { 10D, 14D }, Assert.Single(snapshot.Data.Series).Values);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsForMixedChartJsDatasetTypes() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1", "Q2"],
    "datasets": [
      { "label": "Actual", "data": [3, 7] },
      { "type": "line", "label": "Target", "data": [4, 8] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("mixed per-dataset chart types", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownStackedChartFence_WarnsForSeparateChartJsStackGroups() {
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
    "labels": ["Q1", "Q2"],
    "datasets": [
      { "label": "Actual", "stack": "current", "data": [3, 7] },
      { "label": "Target", "stack": "planned", "data": [4, 8] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("separate stack groups", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsForSecondaryChartJsAxes() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Revenue", "data": [100], "yAxisID": "y" },
      { "label": "Margin", "data": [42], "yAxisID": "y1" }
    ]
  },
  "options": {
    "scales": {
      "y": { "type": "linear" },
      "y1": { "type": "linear" }
    }
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("secondary dataset axes", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_RespectsDisabledChartJsScaleAxes() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1", "Q2"],
    "datasets": [
      { "label": "Actual", "data": [3, 7] }
    ]
  },
  "options": {
    "scales": {
      "x": { "display": false },
      "y": {
        "ticks": { "display": false },
        "grid": { "display": false }
      }
    }
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.False(snapshot!.Layout.ShowCategoryAxis);
        Assert.False(snapshot.Layout.ShowCategoryAxisLine);
        Assert.False(snapshot.Layout.ShowCategoryAxisLabels);
        Assert.True(snapshot.Layout.ShowValueAxis);
        Assert.True(snapshot.Layout.ShowValueAxisLine);
        Assert.False(snapshot.Layout.ShowValueAxisLabels);
        Assert.False(snapshot.Style.ShowGridLines);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_MapsChartJsScaleTitles() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  },
  "options": {
    "scales": {
      "x": { "title": { "display": true, "text": "Quarter" } },
      "y": { "title": { "display": true, "text": "Revenue" } }
    }
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Equal("Quarter", snapshot!.Layout.CategoryAxisTitle);
        Assert.Equal("Revenue", snapshot.Layout.ValueAxisTitle);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_IgnoresChartJsScaleTitleUnlessDisplayIsEnabled() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  },
  "options": {
    "scales": {
      "x": { "title": { "text": "Quarter" } },
      "y": { "title": { "display": false, "text": "Revenue" } }
    }
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Null(snapshot!.Layout.CategoryAxisTitle);
        Assert.Null(snapshot.Layout.ValueAxisTitle);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_RespectsChartJsLegendPosition() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "plugins": {
      "legend": { "position": "top" }
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
        Assert.Equal(OfficeChartLegendPosition.Top, snapshot!.Layout.LegendPosition);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsForUnsupportedChartJsLegendPosition() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "plugins": {
      "legend": { "position": "chartArea" }
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

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("legend position", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsForExplicitChartJsScaleBounds() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "scales": {
      "y": { "min": 5, "max": 20 }
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

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("scale min/max", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsForSuggestedChartJsScaleBounds() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "scales": {
      "y": { "suggestedMin": 0, "suggestedMax": 100 }
    }
  },
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [55] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("suggested bounds", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsForUnsupportedChartJsScaleTypes() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "scales": {
      "y": { "type": "logarithmic" }
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

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("scale types", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsForReversedChartJsScales() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "scales": {
      "y": { "reverse": true }
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

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("reversed scales", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsForLinearCategoryScales() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "options": {
    "scales": {
      "x": { "type": "linear" }
    }
  },
  "data": {
    "datasets": [
      { "label": "Actual", "data": [
        { "x": 10, "y": 2 },
        { "x": 25, "y": 4 }
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("category axis", warning, StringComparison.Ordinal);
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
        Assert.Equal(10D, series.XValues![0]);
        Assert.True(double.IsNaN(series.XValues[1]));
        Assert.Equal(40D, series.XValues[2]);
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
    public void ToPdfDocument_MarkdownChartFence_ExtendsGeneratedLabelsForPositionalDatasetsWhenObjectLabelsExist() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "datasets": [
      { "label": "Object", "data": [
        { "x": "Q1", "y": 1 }
      ] },
      { "label": "Positional", "data": [2, 3] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Equal(new[] { "Q1", "2" }, snapshot!.Data.Categories);
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
    public void ToPdfDocument_MarkdownChartFence_WarnsForFloatingBarTuples() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1", "Q2"],
    "datasets": [
      { "label": "Range", "data": [[2, 5], [4, 8]] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("floating bar tuples", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsForCustomBarBaselines() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Range", "base": 50, "data": [60] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("custom bar baselines", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownScatterChartFence_MarksEarlierScalarPointsMissingWhenExplicitXValuesAppear() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "scatter",
  "data": {
    "datasets": [
      { "label": "Samples", "data": [1000000, { "x": 1, "y": 2 }] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeChartSeries series = Assert.Single(snapshot!.Data.Series);
        Assert.True(double.IsNaN(series.XValues![0]));
        Assert.Equal(1D, series.XValues[1]);
        Assert.True(double.IsNaN(series.Values[0]));
        Assert.Equal(2D, series.Values[1]);
    }

    [Fact]
    public void ToPdfDocument_MarkdownScatterChartFence_DropsExplicitXWhenYIsMissing() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "scatter",
  "data": {
    "datasets": [
      { "label": "Gaps", "data": [
        { "x": 1000000, "y": null },
        { "x": 2, "y": 4 }
      ] },
      { "label": "Visible", "data": [
        { "x": 1, "y": 2 }
      ] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeChartSeries series = snapshot!.Data.Series[0];
        Assert.True(double.IsNaN(series.XValues![0]));
        Assert.True(double.IsNaN(series.Values[0]));
        Assert.Equal(2D, series.XValues[1]);
        Assert.Equal(4D, series.Values[1]);
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
    public void ToPdfDocument_MarkdownScatterChartFence_WarnsForCurvedConnectedLines() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "scatter",
  "options": {
    "showLine": true,
    "elements": { "line": { "tension": 0.35 } }
  },
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

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("stepped or curved", warning, StringComparison.Ordinal);
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
    public void ToPdfDocument_MarkdownChartFence_WarnsForSpanGapsAcrossMissingLinePoints() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "options": {
    "spanGaps": true
  },
  "data": {
    "labels": ["A", "B", "C"],
    "datasets": [
      { "label": "Actual", "data": [1, null, 3] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("spanGaps", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownLineChartFence_DisablesMarkersForChartJsZeroPointRadius() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "options": {
    "elements": { "point": { "radius": 0 } }
  },
  "data": {
    "labels": ["A", "B"],
    "datasets": [
      { "label": "Actual", "data": [1, 3] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.False(snapshot!.Layout.ShowMarkers);
        Assert.False(Assert.Single(snapshot.Data.Series).ShowMarkers);
    }

    [Fact]
    public void ToPdfDocument_MarkdownLineChartFence_WarnsForSteppedOrCurvedLineInterpolation() {
        var stepped = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "data": {
    "labels": ["A", "B"],
    "datasets": [
      { "label": "Actual", "stepped": true, "data": [1, 3] }
    ]
  }
}
""");

        bool steppedCreated = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(stepped, CreateVisualOptions(), out OfficeChartSnapshot? steppedSnapshot, out string? steppedWarning);

        Assert.False(steppedCreated);
        Assert.Null(steppedSnapshot);
        Assert.Contains("stepped or curved", steppedWarning, StringComparison.Ordinal);

        var curved = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "options": {
    "elements": { "line": { "tension": 0.35 } }
  },
  "data": {
    "labels": ["A", "B"],
    "datasets": [
      { "label": "Actual", "data": [1, 3] }
    ]
  }
}
""");

        bool curvedCreated = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(curved, CreateVisualOptions(), out OfficeChartSnapshot? curvedSnapshot, out string? curvedWarning);

        Assert.False(curvedCreated);
        Assert.Null(curvedSnapshot);
        Assert.Contains("stepped or curved", curvedWarning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownLineChartFence_WarnsForDashedLineStyles() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "options": {
    "elements": { "line": { "borderDash": [6, 3] } }
  },
  "data": {
    "labels": ["A", "B"],
    "datasets": [
      { "label": "Actual", "data": [1, 3] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("dashed line styles", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownLineChartFence_WarnsForVerticalOrientation() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "options": {
    "indexAxis": "y"
  },
  "data": {
    "labels": ["A", "B"],
    "datasets": [
      { "label": "Actual", "data": [1, 3] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("vertical line orientation", warning, StringComparison.Ordinal);
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
    public void ToPdfDocument_MarkdownRadarChartFence_ExcludesNonDrawableSeries() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "radar",
  "data": {
    "labels": ["Quality", "Speed", "Cost"],
    "datasets": [
      { "label": "Visible", "data": [1, 2, 3] },
      { "label": "Invisible", "data": [1000, null, null] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeChartSeries series = Assert.Single(snapshot!.Data.Series);
        Assert.Equal("Visible", series.Name);
        Assert.Equal(new[] { 1D, 2D, 3D }, series.Values);
    }

    [Fact]
    public void ToPdfDocument_MarkdownRadarChartFence_ShowsMarkersByDefaultAndHonorsPointRadiusZero() {
        var defaultMarkers = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "radar",
  "data": {
    "labels": ["Quality", "Speed", "Cost"],
    "datasets": [
      { "label": "Score", "data": [8, 9, 7] }
    ]
  }
}
""");

        bool defaultCreated = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(defaultMarkers, CreateVisualOptions(), out OfficeChartSnapshot? defaultSnapshot, out string? defaultWarning);

        Assert.True(defaultCreated, defaultWarning);
        Assert.True(defaultSnapshot!.Layout.ShowMarkers);
        Assert.True(Assert.Single(defaultSnapshot.Data.Series).ShowMarkers);

        var hiddenMarkers = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "radar",
  "options": {
    "elements": { "point": { "radius": 0 } }
  },
  "data": {
    "labels": ["Quality", "Speed", "Cost"],
    "datasets": [
      { "label": "Score", "data": [8, 9, 7] }
    ]
  }
}
""");

        bool hiddenCreated = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(hiddenMarkers, CreateVisualOptions(), out OfficeChartSnapshot? hiddenSnapshot, out string? hiddenWarning);

        Assert.True(hiddenCreated, hiddenWarning);
        Assert.False(hiddenSnapshot!.Layout.ShowMarkers);
        Assert.False(Assert.Single(hiddenSnapshot.Data.Series).ShowMarkers);
    }

    [Fact]
    public void ToPdfDocument_MarkdownRadarChartFence_HonorsDisabledDatasetFill() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "radar",
  "data": {
    "labels": ["Quality", "Speed", "Cost"],
    "datasets": [
      { "label": "Score", "fill": false, "data": [8, 9, 7] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.False(snapshot!.Layout.FillRadarSeries);
    }

    [Fact]
    public void ToPdfDocument_MarkdownRadarChartFence_LeavesDatasetsUnfilledByDefault() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "radar",
  "data": {
    "labels": ["Quality", "Speed", "Cost"],
    "datasets": [
      { "label": "Score", "backgroundColor": "#4472c4", "data": [8, 9, 7] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.False(snapshot!.Layout.FillRadarSeries);
    }

    [Fact]
    public void ToPdfDocument_MarkdownRadarChartFence_WarnsForMixedDefaultAndFilledDatasets() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "radar",
  "data": {
    "labels": ["Quality", "Speed", "Cost"],
    "datasets": [
      { "label": "Score", "fill": true, "data": [8, 9, 7] },
      { "label": "Target", "data": [7, 8, 8] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("filled and unfilled", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownRadarChartFence_WarnsForUnsupportedRadialScaleVisibility() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "radar",
  "options": {
    "scales": {
      "r": { "grid": { "display": false } }
    }
  },
  "data": {
    "labels": ["Quality", "Speed", "Cost"],
    "datasets": [
      { "label": "Score", "data": [8, 9, 7] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("radial scale visibility", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownRadarChartFence_WarnsForUnsupportedPointLabelsVisibility() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "radar",
  "options": {
    "scales": {
      "r": { "pointLabels": { "display": false } }
    }
  },
  "data": {
    "labels": ["Quality", "Speed", "Cost"],
    "datasets": [
      { "label": "Score", "data": [8, 9, 7] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("radial scale visibility", warning, StringComparison.Ordinal);
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
    public void ToPdfDocument_MarkdownAreaChartFence_MapsChartJsStackedAreas() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "area",
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
        Assert.Equal(OfficeChartKind.AreaStacked, snapshot!.ChartKind);
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
    public void ToPdfDocument_MarkdownDoughnutChartFence_WarnsForCustomChartJsCutout() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "doughnut",
  "options": {
    "cutout": "70%"
  },
  "data": {
    "labels": ["Passed", "Failed"],
    "datasets": [
      { "label": "Status", "data": [8, 2] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("custom cutout", warning, StringComparison.Ordinal);
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
    public void ToPdfDocument_MarkdownChartFence_WarnsForTranslucentCssColors() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      {
        "label": "Actual",
        "data": [10],
        "backgroundColor": "rgba(54, 162, 235, 0.2)",
        "borderColor": "#111111"
      }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("translucent colors", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_UsesOpaqueCssBackgroundColorForFilledSeries() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      {
        "label": "Actual",
        "data": [10],
        "backgroundColor": "rgb(54, 162, 235)",
        "borderColor": "#111111"
      }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        OfficeColor color = Assert.Single(snapshot!.Data.Series).Color!.Value;
        Assert.Equal(54, color.R);
        Assert.Equal(162, color.G);
        Assert.Equal(235, color.B);
        Assert.Equal(255, color.A);
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
    public void ToPdfDocument_MarkdownChartFence_RespectsFalseChartJsLegendPluginShorthand() {
        var disabledLegend = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "plugins": {
      "legend": false
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

        bool disabledCreated = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(disabledLegend, CreateVisualOptions(), out OfficeChartSnapshot? disabledSnapshot, out string? disabledWarning);

        Assert.True(disabledCreated, disabledWarning);
        Assert.False(disabledSnapshot!.Layout.ShowLegend);

        var disabledPlugins = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "plugins": false
  },
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  }
}
""");

        bool pluginsCreated = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(disabledPlugins, CreateVisualOptions(), out OfficeChartSnapshot? pluginsSnapshot, out string? pluginsWarning);

        Assert.True(pluginsCreated, pluginsWarning);
        Assert.False(pluginsSnapshot!.Layout.ShowLegend);
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
    public void ToPdfDocument_MarkdownChartFence_RespectsFalseChartJsTitlePluginShorthand() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "plugins": {
      "title": false
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
    public void ToPdfDocument_MarkdownChartFence_MapsFilledChartJsLinesToAreaCharts() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "data": {
    "labels": ["Q1", "Q2"],
    "datasets": [
      { "label": "Actual", "fill": true, "data": [10, 12] },
      { "label": "Forecast", "fill": "origin", "data": [8, 11] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.Equal(OfficeChartKind.Area, snapshot!.ChartKind);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_MapsInheritedFilledChartJsLinesToAreaCharts() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "options": {
    "datasets": {
      "line": { "fill": "origin" }
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
        Assert.Equal(OfficeChartKind.Area, snapshot!.ChartKind);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsForUnsupportedChartJsLineFillTargets() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "data": {
    "labels": ["Q1", "Q2"],
    "datasets": [
      { "label": "Actual", "fill": "+1", "data": [10, 12] },
      { "label": "Forecast", "fill": "+1", "data": [8, 11] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("line fill targets", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_WarnsForMixedFilledAndUnfilledChartJsLines() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "data": {
    "labels": ["Q1", "Q2"],
    "datasets": [
      { "label": "Actual", "fill": true, "data": [10, 12] },
      { "label": "Forecast", "data": [8, 11] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("filled and unfilled", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_EnablesMarkersForChartJsStackedLines() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "line",
  "options": {
    "scales": {
      "y": { "stacked": true }
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
        Assert.Equal(OfficeChartKind.LineStacked, snapshot!.ChartKind);
        Assert.True(snapshot.Layout.ShowMarkers);
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
    public void ToPdfDocument_MarkdownPieChartFence_DoesNotShowDataLabelsUnlessRequested() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "pie",
  "data": {
    "labels": ["Passed", "Failed"],
    "datasets": [
      { "label": "Status", "data": [8, 2] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.False(snapshot!.Layout.ShowDataLabels);
        Assert.False(snapshot.Layout.ShowDataLabelPercentages);
        Assert.False(snapshot.Layout.ShowDataLabelCategoryNames);
    }

    [Fact]
    public void ToPdfDocument_MarkdownCartesianChartFence_ShowsDataLabelsWhenChartJsDataLabelsAreRequested() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "bar",
  "options": {
    "plugins": {
      "datalabels": { "display": true }
    }
  },
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [8] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.True(snapshot!.Layout.ShowDataLabels);
        Assert.True(snapshot.Layout.ShowDataLabelValues);
        Assert.False(snapshot.Layout.ShowDataLabelPercentages);
        Assert.False(snapshot.Layout.ShowDataLabelCategoryNames);
    }

    [Fact]
    public void ToPdfDocument_MarkdownPieChartFence_ShowsDataLabelsWhenChartJsDataLabelsAreRequested() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "pie",
  "options": {
    "plugins": {
      "datalabels": { "display": true }
    }
  },
  "data": {
    "labels": ["Passed", "Failed"],
    "datasets": [
      { "label": "Status", "data": [8, 2] }
    ]
  }
}
""");

        bool created = MarkdownPdfConverterExtensions.TryCreateChartSnapshot(semantic, CreateVisualOptions(), out OfficeChartSnapshot? snapshot, out string? warning);

        Assert.True(created, warning);
        Assert.True(snapshot!.Layout.ShowDataLabels);
        Assert.True(snapshot.Layout.ShowDataLabelPercentages);
        Assert.True(snapshot.Layout.ShowDataLabelCategoryNames);
    }

    [Fact]
    public void ToPdfDocument_MarkdownPieChartFence_WarnsWhenMultipleVisibleDatasetsWouldDropData() {
        var semantic = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "chart", """
{
  "type": "pie",
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

        Assert.False(created);
        Assert.Null(snapshot);
        Assert.Contains("multiple visible datasets", warning, StringComparison.Ordinal);
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
        MarkdownReaderOptions readerOptions = MarkdownReaderOptions.CreateOfficeIMOProfile();
        byte[] bytes = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, readerOptions).ToPdfDocument(options).ToBytes();
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

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

        MarkdownPdfStyle theme = MarkdownPdfStyle.Report();
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
        options.Style = theme;

        byte[] bytes = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdfDocument(options).ToBytes();
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

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
        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, MarkdownPdfSemanticBlocks.CreateReaderOptions()).ToPdfDocumentResult(options);
        byte[] bytes = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains(result.Warnings, warning => warning.Code == "UnsupportedChartFence");
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
        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, MarkdownPdfSemanticBlocks.CreateReaderOptions()).ToPdfDocumentResult(options);
        byte[] bytes = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains(result.Warnings, warning => warning.Code == "UnsupportedSemanticFence" && warning.Source == MarkdownSemanticKinds.Mermaid);
        Assert.Contains("mermaid", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("graph TD", text, StringComparison.Ordinal);
        Assert.Contains("Figure 3. Flow fallback", text, StringComparison.Ordinal);
    }

    private static MarkdownPdfSaveOptions CreateVisualOptions() => new MarkdownPdfSaveOptions {
        Style = MarkdownPdfStyle.Report(),
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

    private static PdfCore.PdfOptions GetPdfOptions(PdfCore.PdfDocument document) {
        PropertyInfo property = typeof(PdfCore.PdfDocument).GetProperty("Options", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (PdfCore.PdfOptions)property.GetValue(document)!;
    }

    private static string CreateDataUriPng() {
        string base64 = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 1));
        return "data:image/png;base64," + base64;
    }

    private static byte[] CreateGifBytes() => Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");
}
