using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Native_Document_Tests {
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
}
