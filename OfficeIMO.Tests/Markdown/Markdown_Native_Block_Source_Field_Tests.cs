using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Native_Block_Source_Field_Tests {
    [Fact]
    public void Footnote_Label_SourceField_Uses_Indented_Label_Token_Span() {
        const string markdown = """
  [^note]: Footnote body
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var footnote = Assert.IsType<MarkdownNativeFootnoteDefinitionBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 8), footnote.LabelSourceSpan);

        var label = Assert.Single(native.EnumerateBlockSourceFields("label"));
        Assert.Same(footnote, label.Block);
        Assert.Equal("note", label.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 8), label.SourceSpan);
        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 6));
        Assert.Equal(label.Name, found.Name);
        Assert.Equal(label.Value, found.Value);
        Assert.Equal(label.SourceSpan, found.SourceSpan);

        var snapshotLabel = Assert.Single(native.ToSnapshot().Blocks[0].SourceFields);
        Assert.Equal("label", snapshotLabel.Name);
        Assert.Equal("note", snapshotLabel.Value);
        Assert.Equal(1, snapshotLabel.SourceSpan.StartLine);
        Assert.Equal(5, snapshotLabel.SourceSpan.StartColumn);
        Assert.Equal(1, snapshotLabel.SourceSpan.EndLine);
        Assert.Equal(8, snapshotLabel.SourceSpan.EndColumn);

        var edited = native.CreateReplaceEdit(label, "memo").Apply(native.SourceMarkdown);
        Assert.Equal("  [^memo]: Footnote body", edited.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void ToSnapshot_Projects_SourceFields_From_The_Same_Native_Block_Field_Enumeration() {
        var markdown = """
# Title

> [!NOTE] Heads up
> Body
> > Nested

<details>
<summary>More</summary>

Inside
</details>

[^note]: Footnote

```cs
Console.WriteLine();
```

| A | B |
| --- | --- |
| 1 | 2 |

> Quote
> Again

---
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var blocks = native.DescendantBlocksAndSelf().ToArray();
        var snapshots = Flatten(native.ToSnapshot().Blocks).ToArray();

        Assert.Equal(blocks.Length, snapshots.Length);

        for (var blockIndex = 0; blockIndex < blocks.Length; blockIndex++) {
            Assert.Equal(blocks[blockIndex].Id, snapshots[blockIndex].Id);
            var fields = MarkdownNativeDocument.EnumerateBlockSourceFields(blocks[blockIndex]).ToArray();

            Assert.Equal(fields.Length, snapshots[blockIndex].SourceFields.Count);
            for (var fieldIndex = 0; fieldIndex < fields.Length; fieldIndex++) {
                var field = fields[fieldIndex];
                var snapshot = snapshots[blockIndex].SourceFields[fieldIndex];

                Assert.Equal(field.Name, snapshot.Name);
                Assert.Equal(field.Value, snapshot.Value);
                Assert.Equal(field.Index, snapshot.Index);
                Assert.Equal(field.SourceSpan.StartLine, snapshot.SourceSpan.StartLine);
                Assert.Equal(field.SourceSpan.StartColumn, snapshot.SourceSpan.StartColumn);
                Assert.Equal(field.SourceSpan.EndLine, snapshot.SourceSpan.EndLine);
                Assert.Equal(field.SourceSpan.EndColumn, snapshot.SourceSpan.EndColumn);
            }
        }

        var allFields = snapshots.SelectMany(block => block.SourceFields).ToArray();
        Assert.Contains(allFields, field => field.Name == "quoteMarker" && field.Index == 0 && field.SourceSpan.StartLine == 23);
        Assert.Contains(allFields, field => field.Name == "quoteMarker" && field.Index == 1 && field.SourceSpan.StartLine == 24);
        Assert.Contains(allFields, field => field.Name == "label" && field.Value == "note");
        Assert.Contains(allFields, field => field.Name == "marker" && field.Value == "---");
    }

    private static IEnumerable<MarkdownNativeBlockSnapshot> Flatten(IReadOnlyList<MarkdownNativeBlockSnapshot> blocks) {
        for (var i = 0; i < blocks.Count; i++) {
            yield return blocks[i];

            foreach (var child in Flatten(blocks[i].Children)) {
                yield return child;
            }

            foreach (var item in blocks[i].Items) {
                foreach (var child in Flatten(item.Children)) {
                    yield return child;
                }
            }

            foreach (var cell in blocks[i].HeaderCells) {
                foreach (var child in Flatten(cell.Children)) {
                    yield return child;
                }
            }

            foreach (var row in blocks[i].Rows) {
                foreach (var cell in row) {
                    foreach (var child in Flatten(cell.Children)) {
                        yield return child;
                    }
                }
            }
        }
    }
}
