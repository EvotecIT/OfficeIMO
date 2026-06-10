using System.Linq;
using OfficeIMO.Markdown;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Native_Document_Tests {
    [Fact]
    public void Parse_Projects_UI_ReadModel_Blocks_With_Stable_Ids_And_Children() {
        var markdown = """
---
title: Native projection
---
# Session

> [!WARNING] Watch
> Body text

> Quoted text

- [x] Done
- Plain

![Signal](images/signal.png "Signal")

<details open>
<summary>More context</summary>

Inside details
</details>
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var reparsed = MarkdownNativeDocument.Parse(markdown);

        Assert.Equal(MarkdownNativeDocumentSourceKind.ReaderInput, native.SourceKind);
        Assert.Equal(markdown, native.SourceMarkdown);
        Assert.Equal(
            native.Blocks.Select(block => block.Id).ToArray(),
            reparsed.Blocks.Select(block => block.Id).ToArray());

        Assert.Equal(new[] {
            MarkdownNativeBlockKind.FrontMatter,
            MarkdownNativeBlockKind.Heading,
            MarkdownNativeBlockKind.Callout,
            MarkdownNativeBlockKind.Quote,
            MarkdownNativeBlockKind.List,
            MarkdownNativeBlockKind.Image,
            MarkdownNativeBlockKind.Details
        }, native.Blocks.Select(block => block.Kind).ToArray());

        var frontMatter = Assert.IsType<MarkdownNativeFrontMatterBlock>(native.Blocks[0]);
        Assert.Equal("Native projection", frontMatter.Values["title"]);

        var heading = Assert.IsType<MarkdownNativeHeadingBlock>(native.Blocks[1]);
        Assert.Equal(1, heading.Level);
        Assert.Equal("Session", heading.Text);

        var callout = Assert.IsType<MarkdownNativeCalloutBlock>(native.Blocks[2]);
        Assert.Equal("warning", callout.CalloutKind);
        Assert.Equal("Watch", callout.Title);
        Assert.Equal("Body text", Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(callout.Children)).Text);
        Assert.Same(callout.Children[0], native.FindBlockAtLine(7));

        var quote = Assert.IsType<MarkdownNativeQuoteBlock>(native.Blocks[3]);
        Assert.Equal("Quoted text", Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(quote.Children)).Text);

        var list = Assert.IsType<MarkdownNativeListBlock>(native.Blocks[4]);
        Assert.False(list.IsOrdered);
        Assert.Equal(2, list.Items.Count);
        Assert.True(list.Items[0].IsTask);
        Assert.True(list.Items[0].Checked);
        Assert.Equal("Done", list.Items[0].Text);
        Assert.NotEmpty(list.Items[0].Id);
        Assert.NotEmpty(list.Items[0].Children);

        var image = Assert.IsType<MarkdownNativeImageBlock>(native.Blocks[5]);
        Assert.Equal("images/signal.png", image.Source);
        Assert.Equal("Signal", image.Alt);
        Assert.Equal("Signal", image.Title);

        var details = Assert.IsType<MarkdownNativeDetailsBlock>(native.Blocks[6]);
        Assert.True(details.Open);
        Assert.Equal("More context", details.Summary);
        Assert.Equal("Inside details", Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(details.Children)).Text);
    }

    [Fact]
    public void Parse_Projects_Core_Blocks_With_SourceSpans() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence));
        var markdown = """
Intro text

```csharp
Console.WriteLine(1);
```

| Name | Value |
| --- | --- |
| CPU | 42 |

```ix-chart
{"type":"bar"}
```
""";

        var native = MarkdownNativeDocument.Parse(markdown, options);

        Assert.Equal(new[] {
            MarkdownNativeBlockKind.Paragraph,
            MarkdownNativeBlockKind.Code,
            MarkdownNativeBlockKind.Table,
            MarkdownNativeBlockKind.Visual
        }, native.Blocks.Select(block => block.Kind).ToArray());

        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(native.Blocks[0]);
        Assert.Equal("Intro text", paragraph.Text);
        Assert.Equal(1, paragraph.SourceSpan!.Value.StartLine);

        var code = Assert.IsType<MarkdownNativeCodeBlock>(native.Blocks[1]);
        Assert.Equal("csharp", code.Language);
        Assert.Equal("Console.WriteLine(1);", code.Content);
        Assert.Equal(3, code.SourceSpan!.Value.StartLine);

        var table = Assert.IsType<MarkdownNativeTableBlock>(native.Blocks[2]);
        Assert.Equal("Name", table.HeaderCells[0].Text);
        Assert.Equal("42", table.Rows[0][1].Text);
        Assert.Equal(7, table.SourceSpan!.Value.StartLine);

        var visual = Assert.IsType<MarkdownNativeVisualBlock>(native.Blocks[3]);
        Assert.Equal(MarkdownSemanticKinds.Chart, visual.SemanticKind);
        Assert.Equal("ix-chart", visual.Language);
        Assert.Equal("{\"type\":\"bar\"}", visual.Content);
        Assert.Equal(11, visual.SourceSpan!.Value.StartLine);
        Assert.Same(visual, native.FindBlockAtLine(12));
    }

    [Fact]
    public void Parse_Does_Not_Project_Phantom_Headers_For_Headerless_Tables() {
        var markdown = """
| One | 1 |
| Two | 2 |
""";

        var native = MarkdownNativeDocument.Parse(markdown);

        var table = Assert.IsType<MarkdownNativeTableBlock>(Assert.Single(native.Blocks));
        Assert.Empty(table.HeaderCells);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal("One", table.Rows[0][0].Text);
        Assert.Equal("2", table.Rows[1][1].Text);
    }

    [Fact]
    public void Parse_Preserves_Table_Column_Alignment_In_Native_Cells() {
        var markdown = """
| Name | Value |
| :--- | ---: |
| CPU | 42 |
""";

        var native = MarkdownNativeDocument.Parse(markdown);

        var table = Assert.IsType<MarkdownNativeTableBlock>(Assert.Single(native.Blocks));
        Assert.Equal(ColumnAlignment.Left, table.HeaderCells[0].Alignment);
        Assert.Equal(ColumnAlignment.Right, table.HeaderCells[1].Alignment);
        Assert.Equal(ColumnAlignment.Left, table.Rows[0][0].Alignment);
        Assert.Equal(ColumnAlignment.Right, table.Rows[0][1].Alignment);
    }

    [Fact]
    public void Parse_Projects_Fence_Metadata_For_Code_And_Visual_Blocks() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence));
        var markdown = """
```csharp {#sample .wide title="Sample Code" copy=false}
Console.WriteLine(1);
```

```ix-chart {#cpu .compact title="CPU Load" pinned=true rows=5}
{"type":"bar"}
```
""";

        var native = MarkdownNativeDocument.Parse(markdown, options);

        var code = Assert.IsType<MarkdownNativeCodeBlock>(native.Blocks[0]);
        Assert.Equal("sample", code.ElementId);
        Assert.Equal("Sample Code", code.Title);
        Assert.Contains("wide", code.Classes);
        Assert.Equal("false", code.Attributes["copy"]);

        var visual = Assert.IsType<MarkdownNativeVisualBlock>(native.Blocks[1]);
        Assert.Equal("cpu", visual.ElementId);
        Assert.Equal("CPU Load", visual.Title);
        Assert.Contains("compact", visual.Classes);
        Assert.Equal("true", visual.Attributes["pinned"]);
        Assert.Equal("5", visual.Attributes["rows"]);
    }

    [Fact]
    public void Renderer_ParseNativeDocument_Exposes_Preprocessed_Source_Kind_And_Transform_Diagnostics() {
        var markdown = """
ix:cached-tool-evidence:v1

```json
{"type":"bar","data":{"labels":["A"],"datasets":[{"label":"Count","data":[1]}]}}
```
""";

        var native = OfficeIMO.MarkdownRenderer.MarkdownRenderer.ParseNativeDocument(
            markdown,
            MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal());

        Assert.Equal(MarkdownNativeDocumentSourceKind.RendererPreprocessed, native.SourceKind);
        Assert.Contains("ix:cached-tool-evidence:v1", native.SourceMarkdown);
        Assert.Contains(native.Diagnostics, diagnostic => diagnostic.Id == "native.transform");
        Assert.Contains(native.Blocks, block =>
            block is MarkdownNativeVisualBlock visual
            && visual.SemanticKind == MarkdownSemanticKinds.Chart
            && visual.Language == "ix-chart");
    }

    [Fact]
    public void Parse_Reports_Fallback_Diagnostics_For_Unsupported_Blocks() {
        var native = MarkdownNativeDocument.Parse("***");

        var other = Assert.IsType<MarkdownNativeOtherBlock>(Assert.Single(native.Blocks));
        var diagnostic = Assert.Single(native.Diagnostics, item => item.Id == "native.unsupported-block");
        Assert.Same(other, diagnostic.Block);
        Assert.Equal(MarkdownNativeDiagnosticSeverity.Info, diagnostic.Severity);
    }
}
