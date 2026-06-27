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

        var snapshot = native.ToSnapshot().Blocks[0];
        var snapshotLabel = Assert.Single(snapshot.SourceFields, field => field.Name == "label");
        Assert.Equal("label", snapshotLabel.Name);
        Assert.Equal("note", snapshotLabel.Value);
        Assert.Equal(1, snapshotLabel.SourceSpan.StartLine);
        Assert.Equal(5, snapshotLabel.SourceSpan.StartColumn);
        Assert.Equal(1, snapshotLabel.SourceSpan.EndLine);
        Assert.Equal(8, snapshotLabel.SourceSpan.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "footnoteBody"
            && field.Value == "Footnote body"
            && field.SourceSpan.StartColumn == 12
            && field.SourceSpan.EndColumn == 24);

        var edited = native.CreateReplaceEdit(label, "memo").Apply(native.SourceMarkdown);
        Assert.Equal("  [^memo]: Footnote body", edited.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void Footnote_Body_SourceField_Uses_Definition_Body_Span() {
        const string markdown = """
[^note]: Old body
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var footnote = Assert.IsType<MarkdownNativeFootnoteDefinitionBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 10, 1, 17), footnote.BodySourceSpan);

        var body = Assert.Single(native.EnumerateBlockSourceFields("footnoteBody"));
        Assert.Same(footnote, body.Block);
        Assert.Equal("Old body", body.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 10, 1, 17), body.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 12));
        Assert.Equal("footnoteBody", found.Name);
        Assert.Equal("Old body", found.Value);
        Assert.Equal(body.SourceSpan, found.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(10, snapshot.FieldSourceSpans["footnoteBody"]!.StartColumn);
        Assert.Equal(17, snapshot.FieldSourceSpans["footnoteBody"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "footnoteBody"
            && field.Value == "Old body"
            && field.SourceSpan.StartColumn == 10
            && field.SourceSpan.EndColumn == 17);

        var edited = native.CreateReplaceEdit(body, "New body").Apply(native.SourceMarkdown);
        Assert.Equal("[^note]: New body", edited.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void Container_Body_SourceFields_Use_Structured_Child_Block_Spans() {
        var markdown = """
> [!TIP] Title
> Body line

<details>
<summary>More</summary>

Inside
</details>
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var callout = Assert.IsType<MarkdownNativeCalloutBlock>(native.Blocks[0]);
        var details = Assert.IsType<MarkdownNativeDetailsBlock>(native.Blocks[1]);

        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 11), callout.BodySourceSpan);
        Assert.Equal(new MarkdownSourceSpan(7, 1, 7, 6), details.BodySourceSpan);

        var calloutBody = Assert.Single(native.EnumerateBlockSourceFields("calloutBody"));
        Assert.Same(callout, calloutBody.Block);
        Assert.Equal("Body line", calloutBody.Value);
        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 11), calloutBody.SourceSpan);

        var detailsBody = Assert.Single(native.EnumerateBlockSourceFields("detailsBody"));
        Assert.Same(details, detailsBody.Block);
        Assert.Null(detailsBody.Value);
        Assert.Equal(new MarkdownSourceSpan(7, 1, 7, 6), detailsBody.SourceSpan);

        Assert.Equal("calloutBody", native.FindBlockSourceFieldAtPosition(2, 4)!.Name);
        Assert.Equal("detailsBody", native.FindBlockSourceFieldAtPosition(7, 3)!.Name);

        var snapshot = native.ToSnapshot();
        Assert.Equal(3, snapshot.Blocks[0].FieldSourceSpans["calloutBody"]!.StartColumn);
        Assert.Equal(11, snapshot.Blocks[0].FieldSourceSpans["calloutBody"]!.EndColumn);
        Assert.Equal(1, snapshot.Blocks[1].FieldSourceSpans["detailsBody"]!.StartColumn);
        Assert.Equal(6, snapshot.Blocks[1].FieldSourceSpans["detailsBody"]!.EndColumn);
        Assert.Contains(snapshot.Blocks[0].SourceFields, field => field.Name == "calloutBody" && field.Value == "Body line");
        Assert.Contains(snapshot.Blocks[1].SourceFields, field => field.Name == "detailsBody" && field.Value == null);

        Assert.Contains("> Updated", native.CreateReplaceEdit(calloutBody, "Updated").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("Outside", native.CreateReplaceEdit(detailsBody, "Outside").Apply(native.SourceMarkdown), StringComparison.Ordinal);
    }

    [Fact]
    public void Quote_Body_SourceField_Uses_Structured_Child_Block_Span() {
        const string markdown = "> Old quote\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var quote = Assert.IsType<MarkdownNativeQuoteBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 11), quote.BodySourceSpan);

        var body = Assert.Single(native.EnumerateBlockSourceFields("quoteBody"));
        Assert.Same(quote, body.Block);
        Assert.Null(body.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 11), body.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 5));
        Assert.Equal("quoteBody", found.Name);
        Assert.Equal(body.SourceSpan, found.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(3, snapshot.FieldSourceSpans["quoteBody"]!.StartColumn);
        Assert.Equal(11, snapshot.FieldSourceSpans["quoteBody"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "quoteBody"
            && field.Value == null
            && field.SourceSpan.StartColumn == 3
            && field.SourceSpan.EndColumn == 11);

        var edited = native.CreateReplaceEdit(body, "New quote").Apply(native.SourceMarkdown);
        Assert.Equal("> New quote", edited.TrimEnd('\r', '\n'));
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
        Assert.Contains(allFields, field => field.Name == "quoteBody" && field.SourceSpan.StartLine == 23);
        Assert.Contains(allFields, field => field.Name == "calloutBody" && field.Value!.Replace("\r\n", "\n") == "Body\n\n> Nested");
        Assert.Contains(allFields, field => field.Name == "detailsBody");
        Assert.Contains(allFields, field => field.Name == "label" && field.Value == "note");
        Assert.Contains(allFields, field => field.Name == "footnoteBody" && field.Value == "Footnote");
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
