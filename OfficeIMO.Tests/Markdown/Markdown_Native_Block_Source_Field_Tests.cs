using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Native_Block_Source_Field_Tests {
    [Fact]
    public void Paragraph_Text_SourceField_Uses_Paragraph_Source_Span() {
        const string markdown = "Old **body**\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 12), paragraph.TextSourceSpan);

        var text = Assert.Single(native.EnumerateBlockSourceFields("paragraphText"));
        Assert.Same(paragraph, text.Block);
        Assert.Equal("Old body", text.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 12), text.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 6));
        Assert.Equal("paragraphText", found.Name);
        Assert.Equal(text.Value, found.Value);
        Assert.Equal(text.SourceSpan, found.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(1, snapshot.FieldSourceSpans["paragraphText"]!.StartColumn);
        Assert.Equal(12, snapshot.FieldSourceSpans["paragraphText"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "paragraphText"
            && field.Value == "Old body"
            && field.SourceSpan.StartColumn == 1
            && field.SourceSpan.EndColumn == 12);

        var edited = native.CreateReplaceEdit(text, "New _body_").Apply(native.SourceMarkdown);
        Assert.Equal("New _body_", edited.TrimEnd('\r', '\n'));
    }

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
    public void Image_SourceFields_Use_Image_Token_Spans() {
        const string markdown = "![Alt text](https://example.com/image.png \"Image title\")\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var image = Assert.IsType<MarkdownNativeImageBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 10), image.AltSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 13, 1, 41), image.SourceSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 44, 1, 54), image.TitleSourceSpan);

        var alt = Assert.Single(native.EnumerateBlockSourceFields("alt"));
        Assert.Same(image, alt.Block);
        Assert.Equal("Alt text", alt.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 10), alt.SourceSpan);

        var source = Assert.Single(native.EnumerateBlockSourceFields("source"));
        Assert.Equal("https://example.com/image.png", source.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 13, 1, 41), source.SourceSpan);

        var title = Assert.Single(native.EnumerateBlockSourceFields("title"));
        Assert.Equal("Image title", title.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 44, 1, 54), title.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 45));
        Assert.Equal("title", found.Name);
        Assert.Equal("Image title", found.Value);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(3, snapshot.FieldSourceSpans["alt"]!.StartColumn);
        Assert.Equal(41, snapshot.FieldSourceSpans["source"]!.EndColumn);
        Assert.Equal(44, snapshot.FieldSourceSpans["title"]!.StartColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "source"
            && field.Value == "https://example.com/image.png"
            && field.SourceSpan.StartColumn == 13
            && field.SourceSpan.EndColumn == 41);

        var edited = native.CreateReplaceEdit(source, "/media/new.png").Apply(native.SourceMarkdown);
        Assert.Equal("![Alt text](/media/new.png \"Image title\")", edited.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void Linked_Image_SourceFields_Use_Image_And_Link_Token_Spans() {
        var markdown = """
[![Alt text](https://example.com/image.png "Image title")](https://example.com/docs "Link title")
_Caption_
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var image = Assert.IsType<MarkdownNativeImageBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 4, 1, 11), image.AltSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 14, 1, 42), image.SourceSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 45, 1, 55), image.TitleSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 60, 1, 83), image.LinkUrlSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 86, 1, 95), image.LinkTitleSourceSpan);

        var linkUrl = Assert.Single(native.EnumerateBlockSourceFields("linkUrl"));
        Assert.Same(image, linkUrl.Block);
        Assert.Equal("https://example.com/docs", linkUrl.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 60, 1, 83), linkUrl.SourceSpan);

        var linkTitle = Assert.Single(native.EnumerateBlockSourceFields("linkTitle"));
        Assert.Equal("Link title", linkTitle.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 86, 1, 95), linkTitle.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 62));
        Assert.Equal("linkUrl", found.Name);
        Assert.Equal("https://example.com/docs", found.Value);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(60, snapshot.FieldSourceSpans["linkUrl"]!.StartColumn);
        Assert.Equal(95, snapshot.FieldSourceSpans["linkTitle"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "linkTitle"
            && field.Value == "Link title"
            && field.SourceSpan.StartColumn == 86
            && field.SourceSpan.EndColumn == 95);

        var edited = native.CreateReplaceEdit(linkTitle, "More docs").Apply(native.SourceMarkdown);
        Assert.Contains("](https://example.com/docs \"More docs\")", edited, StringComparison.Ordinal);
        Assert.Contains("_Caption_", edited, StringComparison.Ordinal);
    }

    [Fact]
    public void FrontMatter_SourceFields_Use_Entry_Key_And_Value_Token_Spans() {
        var markdown = """
---
title: Doc
published: true
tags: [a, b]
summary: |
  First line
  Second line
---

# Heading
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var frontMatter = Assert.IsType<MarkdownNativeFrontMatterBlock>(native.Blocks[0]);

        Assert.Collection(frontMatter.Entries,
            entry => {
                Assert.Equal("title", entry.Key);
                Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 5), entry.KeySourceSpan);
                Assert.Equal(new MarkdownSourceSpan(2, 8, 2, 10), entry.ValueSourceSpan);
            },
            entry => {
                Assert.Equal("published", entry.Key);
                Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 9), entry.KeySourceSpan);
                Assert.Equal(new MarkdownSourceSpan(3, 12, 3, 15), entry.ValueSourceSpan);
            },
            entry => {
                Assert.Equal("tags", entry.Key);
                Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 4), entry.KeySourceSpan);
                Assert.Equal(new MarkdownSourceSpan(4, 7, 4, 12), entry.ValueSourceSpan);
            },
            entry => {
                Assert.Equal("summary", entry.Key);
                Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 7), entry.KeySourceSpan);
                Assert.Equal(new MarkdownSourceSpan(6, 3, 7, 13), entry.ValueSourceSpan);
            });

        var keys = native.EnumerateBlockSourceFields("frontMatterKey").ToArray();
        Assert.Equal(4, keys.Length);
        Assert.Equal("published", keys[1].Value);
        Assert.Equal(1, keys[1].Index);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 9), keys[1].SourceSpan);

        var values = native.EnumerateBlockSourceFields("frontMatterValue").ToArray();
        Assert.Equal(4, values.Length);
        Assert.Equal("true", values[1].Value);
        Assert.Equal("a, b", values[2].Value);
        Assert.Equal("First line\nSecond line", values[3].Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(6, 3, 7, 13), values[3].SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(4, 8));
        Assert.Equal("frontMatterValue", found.Name);
        Assert.Equal(2, found.Index);
        Assert.Equal("a, b", found.Value);

        var snapshot = native.ToSnapshot().Blocks[0];
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "frontMatterKey"
            && field.Value == "summary"
            && field.Index == 3
            && field.SourceSpan.StartLine == 5
            && field.SourceSpan.EndColumn == 7);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "frontMatterValue"
            && field.Value!.Replace("\r\n", "\n") == "First line\nSecond line"
            && field.Index == 3
            && field.SourceSpan.StartLine == 6
            && field.SourceSpan.StartColumn == 3);

        var edited = native.CreateReplaceEdit(values[0], "Better").Apply(native.SourceMarkdown);
        Assert.Contains("title: Better", edited, StringComparison.Ordinal);
        Assert.Contains("published: true", edited, StringComparison.Ordinal);
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

![Alt text](https://example.com/image.png "Image title")

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
        Assert.Contains(allFields, field => field.Name == "paragraphText" && field.Value == "Body");
        Assert.Contains(allFields, field => field.Name == "quoteMarker" && field.Index == 0 && field.SourceSpan.StartLine == 23);
        Assert.Contains(allFields, field => field.Name == "quoteMarker" && field.Index == 1 && field.SourceSpan.StartLine == 24);
        Assert.Contains(allFields, field => field.Name == "quoteBody" && field.SourceSpan.StartLine == 23);
        Assert.Contains(allFields, field => field.Name == "calloutBody" && field.Value!.Replace("\r\n", "\n") == "Body\n\n> Nested");
        Assert.Contains(allFields, field => field.Name == "detailsBody");
        Assert.Contains(allFields, field => field.Name == "label" && field.Value == "note");
        Assert.Contains(allFields, field => field.Name == "footnoteBody" && field.Value == "Footnote");
        Assert.Contains(allFields, field => field.Name == "marker" && field.Value == "---");
        Assert.Contains(allFields, field => field.Name == "source" && field.Value == "https://example.com/image.png");
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
