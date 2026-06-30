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
        Assert.Equal(NormalizeLineEndings(markdown), native.SourceMarkdown);
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
        Assert.Equal(new MarkdownSourceSpan(6, 3, 6, 4), callout.OpeningMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(6, 5, 6, 11), callout.KindSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(6, 12, 6, 12), callout.ClosingMarkerSourceSpan);
        Assert.Equal("Watch", callout.Title);
        Assert.Equal(new MarkdownSourceSpan(6, 14, 6, 18), callout.TitleSourceSpan);
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
        Assert.Equal(new MarkdownSourceSpan(14, 3, 14, 8), image.AltSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(14, 11, 14, 27), image.SourceSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(14, 30, 14, 35), image.TitleSourceSpan);

        var details = Assert.IsType<MarkdownNativeDetailsBlock>(native.Blocks[6]);
        Assert.True(details.Open);
        Assert.Equal("More context", details.Summary);
        Assert.Equal(new MarkdownSourceSpan(17, 17), details.SummarySourceSpan);
        Assert.Equal("Inside details", Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(details.Children)).Text);

        var detailsSnapshot = native.ToSnapshot().Blocks[6];
        Assert.Equal(17, detailsSnapshot.FieldSourceSpans["summary"]!.StartLine);
        Assert.Equal(17, detailsSnapshot.FieldSourceSpans["summary"]!.EndLine);

        var imageSnapshot = native.ToSnapshot().Blocks[5];
        Assert.Equal(11, imageSnapshot.FieldSourceSpans["source"]!.StartColumn);
        Assert.Equal(27, imageSnapshot.FieldSourceSpans["source"]!.EndColumn);

        var withSummary = native.CreateReplaceEdit(details.SummarySourceSpan!.Value, "<summary>Less context</summary>").Apply(native.SourceMarkdown);
        Assert.Contains("<summary>Less context</summary>", withSummary);
        Assert.DoesNotContain("<summary>More context</summary>", withSummary);
    }

    [Fact]
    public void Parse_Projects_Task_List_Marker_SourceSpans_Into_Native_Snapshots_And_Edits() {
        const string markdown = "- [X]\tUpper\n- [ ]    Open\n- [x]tight\n";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var list = Assert.IsType<MarkdownNativeListBlock>(Assert.Single(native.Blocks));

        Assert.Collection(
            list.Items,
            item => {
                Assert.True(item.IsTask);
                Assert.True(item.Checked);
                Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 1), item.MarkerSourceSpan);
                Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 5), item.TaskMarkerSourceSpan);
            },
            item => {
                Assert.True(item.IsTask);
                Assert.False(item.Checked);
                Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 1), item.MarkerSourceSpan);
                Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 5), item.TaskMarkerSourceSpan);
            },
            item => {
                Assert.False(item.IsTask);
                Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 1), item.MarkerSourceSpan);
                Assert.Null(item.TaskMarkerSourceSpan);
            });

        var snapshot = native.ToSnapshot();
        var snapshotList = Assert.Single(snapshot.Blocks);
        Assert.Equal(1, snapshotList.Items[0].MarkerSourceSpan!.StartColumn);
        Assert.Equal(1, snapshotList.Items[0].MarkerSourceSpan!.EndColumn);
        Assert.Equal(3, snapshotList.Items[0].TaskMarkerSourceSpan!.StartColumn);
        Assert.Equal(5, snapshotList.Items[0].TaskMarkerSourceSpan!.EndColumn);
        Assert.Null(snapshotList.Items[2].TaskMarkerSourceSpan);

        var edited = native.CreateReplaceEdit(list.Items[0].TaskMarkerSourceSpan!.Value, "[ ]").Apply(native.SourceMarkdown);
        Assert.StartsWith("- [ ]\tUpper", edited, StringComparison.Ordinal);
        Assert.Contains("- [ ]    Open", edited, StringComparison.Ordinal);
        Assert.Contains("- [x]tight", edited, StringComparison.Ordinal);

        var withSecondMarker = native.CreateReplaceEdit(list.Items[1].MarkerSourceSpan!.Value, "*").Apply(native.SourceMarkdown);
        Assert.Contains("* [ ]    Open", withSecondMarker, StringComparison.Ordinal);
        Assert.StartsWith("- [X]\tUpper", withSecondMarker, StringComparison.Ordinal);
    }

    [Fact]
    public void Parse_Projects_Quote_Marker_SourceSpans_Into_Native_Snapshots_And_Edits() {
        var markdown = """
> Outer
>
> > Inner
>
> Still outer

After
""";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var quote = Assert.IsType<MarkdownNativeQuoteBlock>(native.Blocks[0]);

        Assert.Equal(new[] {
            new MarkdownSourceSpan(1, 1, 1, 1),
            new MarkdownSourceSpan(2, 1, 2, 1),
            new MarkdownSourceSpan(3, 1, 3, 1),
            new MarkdownSourceSpan(4, 1, 4, 1),
            new MarkdownSourceSpan(5, 1, 5, 1)
        }, quote.MarkerSourceSpans);
        Assert.Equal(quote.MarkerSourceSpans, quote.Quote.MarkerSourceSpans);

        var innerQuote = Assert.IsType<MarkdownNativeQuoteBlock>(quote.Children[1]);
        var innerMarker = Assert.Single(innerQuote.MarkerSourceSpans);
        Assert.Equal(new MarkdownSourceSpan(3, 3, 3, 3), innerMarker);

        var snapshot = native.ToSnapshot().Blocks[0];
        Assert.Equal(5, snapshot.MarkerSourceSpans.Count);
        Assert.Equal(1, snapshot.MarkerSourceSpans[0].StartLine);
        Assert.Equal(1, snapshot.MarkerSourceSpans[0].StartColumn);
        Assert.Equal(3, snapshot.Children[1].MarkerSourceSpans[0].StartLine);
        Assert.Equal(3, snapshot.Children[1].MarkerSourceSpans[0].StartColumn);

        var edited = native.CreateReplaceEdit(quote.MarkerSourceSpans[4], ">>").Apply(native.SourceMarkdown);
        Assert.Contains(">> Still outer", edited, StringComparison.Ordinal);
        Assert.Contains("> > Inner", edited, StringComparison.Ordinal);
    }

    [Fact]
    public void Parse_Projects_Nested_Quote_And_NonOne_Ordered_List_SourceSpans_Into_Native_Snapshots_And_Edits() {
        var markdown = """
- outer
  > alpha
  10. beta
      gamma
""";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var list = Assert.IsType<MarkdownNativeListBlock>(Assert.Single(native.Blocks));
        var item = Assert.Single(list.Items);
        var quote = Assert.Single(item.Children.OfType<MarkdownNativeQuoteBlock>());
        var ordered = Assert.Single(item.Children.OfType<MarkdownNativeListBlock>());
        var orderedItem = Assert.Single(ordered.Items);

        Assert.False(list.IsOrdered);
        Assert.Equal("outer", item.Text);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 1), item.MarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 3), Assert.Single(quote.MarkerSourceSpans));
        Assert.True(ordered.IsOrdered);
        Assert.Equal(10, ordered.Start);
        Assert.Equal("beta gamma", orderedItem.Text);
        Assert.Equal(new MarkdownSourceSpan(3, 3, 3, 5), orderedItem.MarkerSourceSpan);
        Assert.Same(quote, native.FindBlockAtPosition(2, 3));
        Assert.Same(ordered, native.FindBlockAtPosition(3, 3));

        var snapshot = native.ToSnapshot();
        var snapshotItem = Assert.Single(Assert.Single(snapshot.Blocks).Items);
        Assert.Equal(1, snapshotItem.MarkerSourceSpan!.StartColumn);
        var snapshotQuote = Assert.Single(snapshotItem.Children, child => child.Kind == MarkdownNativeBlockKind.Quote);
        Assert.Equal(3, snapshotQuote.MarkerSourceSpans[0].StartColumn);
        var snapshotOrdered = Assert.Single(snapshotItem.Children, child =>
            child.Kind == MarkdownNativeBlockKind.List &&
            string.Equals(child.Fields["isOrdered"], "true", StringComparison.Ordinal) &&
            string.Equals(child.Fields["start"], "10", StringComparison.Ordinal));
        Assert.Equal(3, snapshotOrdered.Items[0].MarkerSourceSpan!.StartColumn);
        Assert.Equal(5, snapshotOrdered.Items[0].MarkerSourceSpan!.EndColumn);

        var withOuterMarker = native.CreateReplaceEdit(item.MarkerSourceSpan!.Value, "*").Apply(native.SourceMarkdown);
        Assert.StartsWith("* outer", withOuterMarker, StringComparison.Ordinal);
        Assert.Contains("  10. beta", withOuterMarker, StringComparison.Ordinal);

        var withOrderedMarker = native.CreateReplaceEdit(orderedItem.MarkerSourceSpan!.Value, "11.").Apply(native.SourceMarkdown);
        Assert.Contains("  11. beta", withOrderedMarker, StringComparison.Ordinal);
        Assert.Contains("  > alpha", withOrderedMarker, StringComparison.Ordinal);
    }

    [Fact]
    public void Source_Edit_Helpers_Replace_Native_List_Item_Content_And_Preserve_Surrounding_Source() {
        var markdown = """
- outer
  - old
  - keep

After
""";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var list = Assert.IsType<MarkdownNativeListBlock>(native.Blocks[0]);
        var outer = Assert.Single(list.Items);
        var nested = Assert.Single(outer.Children.OfType<MarkdownNativeListBlock>());
        var oldItem = nested.Items[0];
        Assert.Equal(new MarkdownSourceSpan(2, 5, 2, 7), oldItem.ContentSourceSpan);

        var edited = native.CreateReplaceEdit(oldItem, "new\n    continuation").Apply(native.SourceMarkdown);

        Assert.Equal(
            """
- outer
  - new
    continuation
  - keep

After
""",
            edited);
    }

    [Fact]
    public void Source_Edit_Helpers_Roundtrip_Native_List_Item_Content_Edits_Against_Original_Source() {
        const string markdown = "- outer\r\n  - old\r\n  - keep\r\n\r\nAfter\r\n";
        var native = MarkdownNativeDocument.Parse(
            markdown,
            new MarkdownReaderOptions { PreserveTrivia = true });
        var list = Assert.IsType<MarkdownNativeListBlock>(native.Blocks[0]);
        var outer = Assert.Single(list.Items);
        var nested = Assert.Single(outer.Children.OfType<MarkdownNativeListBlock>());
        var oldItem = nested.Items[0];
        Assert.Equal(new MarkdownSourceSpan(2, 5, 2, 7), oldItem.ContentSourceSpan);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(oldItem, "new\r\n    continuation"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("- outer\r\n  - new\r\n    continuation\r\n  - keep\r\n\r\nAfter\r\n", roundtrip.Markdown);
    }

    [Fact]
    public void Navigation_Helpers_Find_Native_List_Items_By_Position_And_Id() {
        var markdown = """
- [x] done
  - child
  - keep
""";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var items = native.EnumerateListItems().ToArray();

        Assert.Equal(3, items.Length);
        Assert.Same(items[0], native.FindListItemAtPosition(1, 1));
        Assert.Same(items[0], native.FindListItemAtPosition(1, 4));
        Assert.Same(items[0], native.FindListItemAtPosition(1, 7));
        Assert.Same(items[1], native.FindListItemAtPosition(2, 5));
        Assert.Same(items[1], native.FindListItemById(items[1].Id));
        Assert.Null(native.FindListItemAtPosition(4, 1));
        Assert.Null(native.FindListItemById("missing"));
    }

    [Fact]
    public void Navigation_Helpers_Find_Native_Table_Cells_By_Position() {
        var markdown = """
| Name | Value |
| --- | --- |
| CPU | 42 |
""";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var cells = native.EnumerateTableCells().ToArray();

        Assert.Equal(4, cells.Length);
        Assert.True(cells[1].IsHeader);
        Assert.False(cells[3].IsHeader);
        Assert.Equal(0, cells[3].RowIndex);
        Assert.Equal(1, cells[3].ColumnIndex);
        Assert.Same(cells[1], native.FindTableCellAtPosition(1, 10));
        Assert.Same(cells[3], native.FindTableCellAtPosition(3, 9));
        Assert.Null(native.FindTableCellAtPosition(2, 3));
    }

    [Fact]
    public void Native_Table_Cell_Navigation_Snapshot_And_Source_Edit_Cover_List_Nested_Tables() {
        const string markdown = "- item\r\n\r\n  | Name | Value |\r\n  | --- | --- |\r\n  | CPU | 42 |\r\n";
        var native = MarkdownNativeDocument.Parse(
            markdown,
            new MarkdownReaderOptions {
                PreserveTrivia = true,
                Tables = true
            });

        var cells = native.EnumerateTableCells().ToArray();

        Assert.Equal(4, cells.Length);
        Assert.True(cells[0].IsHeader);
        Assert.False(cells[3].IsHeader);
        Assert.Equal("42", cells[3].Text);
        Assert.Equal(new MarkdownSourceSpan(5, 11, 5, 12), cells[3].SourceSpan);
        Assert.Same(cells[3], native.FindTableCellAtPosition(5, 11));
        Assert.Same(cells[3], native.FindTableCellAtPosition(5, 12));

        var snapshot = native.ToSnapshot();
        var list = Assert.Single(snapshot.Blocks);
        var item = Assert.Single(list.Items);
        Assert.Equal(2, item.Children.Count);
        Assert.Contains(item.Children, child => child.Kind == MarkdownNativeBlockKind.Paragraph);
        var nestedTable = Assert.Single(item.Children, child => child.Kind == MarkdownNativeBlockKind.Table);
        Assert.Equal(MarkdownNativeBlockKind.Table, nestedTable.Kind);
        Assert.Equal("42", nestedTable.Rows[0][1].Text);
        Assert.Equal(5, nestedTable.Rows[0][1].SourceSpan!.StartLine);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(cells[3], "84"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("- item\r\n\r\n  | Name | Value |\r\n  | --- | --- |\r\n  | CPU | 84 |\r\n", roundtrip.Markdown);
    }

    [Fact]
    public void Navigation_Helpers_Find_Native_Definition_List_Parts_By_Position() {
        var markdown = """
First: Intro

  - child

Second: Other
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var groups = native.EnumerateDefinitionListGroups().ToArray();
        var terms = native.EnumerateDefinitionListTerms().ToArray();
        var definitions = native.EnumerateDefinitionListDefinitions().ToArray();

        Assert.Equal(2, groups.Length);
        Assert.Equal(new[] { "First", "Second" }, terms.Select(term => term.Text).ToArray());
        Assert.Equal(new[] { "Intro\n\n- child", "Other" }, definitions.Select(definition => definition.Markdown.Replace("\r\n", "\n")).ToArray());

        Assert.Same(groups[0], native.FindDefinitionListGroupAtPosition(1, 1));
        Assert.Same(groups[0], native.FindDefinitionListGroupAtPosition(3, 5));
        Assert.Same(terms[0], native.FindDefinitionListTermAtPosition(1, 2));
        Assert.Same(definitions[0], native.FindDefinitionListDefinitionAtPosition(1, 8));
        Assert.Same(definitions[0], native.FindDefinitionListDefinitionAtPosition(3, 5));
        Assert.Same(groups[1], native.FindDefinitionListGroupAtPosition(5, 1));
        Assert.Same(terms[1], native.FindDefinitionListTermAtPosition(5, 2));
        Assert.Same(definitions[1], native.FindDefinitionListDefinitionAtPosition(5, 9));
        Assert.Null(native.FindDefinitionListGroupAtPosition(6, 1));
        Assert.Null(native.FindDefinitionListTermAtPosition(3, 5));
        Assert.Null(native.FindDefinitionListDefinitionAtPosition(1, 2));
    }

    [Fact]
    public void Navigation_Helpers_Find_Native_Inline_Metadata_By_Position_And_Name() {
        var markdown = """
# Native [docs](https://example.com "Docs")

Paragraph with footnote[^note] and ![Alt](img.png "Img").
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var target = Assert.Single(native.EnumerateInlineMetadata("TARGET"));
        var title = Assert.Single(native.EnumerateInlineMetadata("title"));
        var label = Assert.Single(native.EnumerateInlineMetadata("label"));
        var source = Assert.Single(native.EnumerateInlineMetadata("source"));

        Assert.Equal("https://example.com", target.Value);
        Assert.Equal("Docs", title.Value);
        Assert.Equal("note", label.Value);
        Assert.Equal("img.png", source.Value);
        Assert.Contains(native.EnumerateInlineMetadata(), metadata => metadata.Name == "alt" && metadata.Value == "Alt");
        Assert.Empty(native.EnumerateInlineMetadata("missing"));
        Assert.Same(target, native.FindInlineMetadataAtPosition(1, 18));
        Assert.Same(title, native.FindInlineMetadataAtPosition(1, 39));
        Assert.Same(label, native.FindInlineMetadataAtPosition(3, 26));
        Assert.Same(source, native.FindInlineMetadataAtPosition(3, 44));
        Assert.Null(native.FindInlineMetadataAtPosition(2, 1));

        var snapshot = native.ToSnapshot();
        var linkSnapshot = Assert.Single(snapshot.Blocks[0].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var linkMetadata = linkSnapshot.MetadataFields.ToArray();
        Assert.Equal(new[] { "openingMarker", "separatorMarker", "target", "title", "closingMarker" }, linkMetadata.Select(field => field.Name).ToArray());
        Assert.Equal("https://example.com", linkSnapshot.FindMetadataField("target")!.Value);
        Assert.Equal("Docs", Assert.Single(linkSnapshot.EnumerateMetadataFields("title")).Value);
        Assert.Equal(2, linkSnapshot.FindMetadataField("target")!.Index);
        Assert.Equal(3, linkSnapshot.FindMetadataField("title")!.Index);
        Assert.NotNull(linkSnapshot.FindMetadataField("target")!.SourceSpan);
        Assert.NotNull(linkSnapshot.FindMetadataField("title")!.SourceSpan);

        var paragraphSnapshot = snapshot.Blocks[1];
        var imageSnapshot = Assert.Single(paragraphSnapshot.Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Image);
        var imageMetadata = imageSnapshot.MetadataFields.ToArray();
        Assert.Equal(new[] { "openingMarker", "alt", "separatorMarker", "source", "imageTitle", "closingMarker" }, imageMetadata.Select(field => field.Name).ToArray());
        Assert.Equal("![", imageSnapshot.FindMetadataField("openingMarker")!.Value);
        Assert.Equal("](", imageSnapshot.FindMetadataField("separatorMarker")!.Value);
        Assert.Equal(")", imageSnapshot.FindMetadataField("closingMarker")!.Value);
        Assert.Equal(36, imageSnapshot.FindMetadataField("openingMarker")!.SourceSpan!.StartColumn);
        Assert.Equal(42, imageSnapshot.FindMetadataField("separatorMarker")!.SourceSpan!.EndColumn);
        Assert.Equal(56, imageSnapshot.FindMetadataField("closingMarker")!.SourceSpan!.StartColumn);

        var imageOpening = Assert.Single(native.EnumerateInlineMetadata("openingMarker"), metadata => metadata.Value == "![" && metadata.SourceSpan?.StartLine == 3);
        var imageSeparator = Assert.Single(native.EnumerateInlineMetadata("separatorMarker"), metadata => metadata.Value == "](" && metadata.SourceSpan?.StartLine == 3);
        var imageClosing = Assert.Single(native.EnumerateInlineMetadata("closingMarker"), metadata => metadata.Value == ")" && metadata.SourceSpan?.StartLine == 3);
        Assert.Contains("and [Alt](img.png \"Img\").", native.CreateReplaceEdit(imageOpening, "[").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("![Alt]][img.png \"Img\").", native.CreateReplaceEdit(imageSeparator, "]][").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("![Alt](img.png \"Img\"]", native.CreateReplaceEdit(imageClosing, "]").Apply(native.SourceMarkdown), StringComparison.Ordinal);
    }

    [Fact]
    public void Native_Inline_Metadata_Snapshots_Preserve_Source_Token_Order_For_Escapes_Entities_And_Breaks() {
        var markdown = "Escaped \\* and &copy;  \nnext";

        var native = MarkdownNativeDocument.Parse(markdown);

        var escapeMarker = Assert.Single(native.EnumerateInlineMetadata("escapeMarker"));
        var escapedCharacter = Assert.Single(native.EnumerateInlineMetadata("escapedCharacter"));
        var entitySourceText = Assert.Single(native.EnumerateInlineMetadata("sourceText"));
        var hardBreakMarker = Assert.Single(native.EnumerateInlineMetadata("marker"));
        Assert.Equal("\\", escapeMarker.Value);
        Assert.Equal("*", escapedCharacter.Value);
        Assert.Equal("&copy;", entitySourceText.Value);
        Assert.Equal("  ", hardBreakMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 9, 1, 9), escapeMarker.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 10, 1, 10), escapedCharacter.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 16, 1, 21), entitySourceText.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 22, 1, 23), hardBreakMarker.SourceSpan);

        var snapshot = native.ToSnapshot();
        Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        Assert.Equal(MarkdownNativeBlockKind.Paragraph, snapshot.Blocks[0].Kind);
        var inlineFields = snapshot.Blocks[0].Inlines
            .SelectMany(inline => inline.MetadataFields)
            .OrderBy(field => field.SourceSpan?.StartLine ?? int.MaxValue)
            .ThenBy(field => field.SourceSpan?.StartColumn ?? int.MaxValue)
            .ToArray();
        Assert.Equal(new[] { "escapeMarker", "escapedCharacter", "sourceText", "marker" }, inlineFields.Select(field => field.Name).ToArray());
        Assert.Equal(new[] { "\\", "*", "&copy;", "  " }, inlineFields.Select(field => field.Value).ToArray());
        Assert.Equal(9, inlineFields[0].SourceSpan!.StartColumn);
        Assert.Equal(23, inlineFields[3].SourceSpan!.EndColumn);

        Assert.Equal("Escaped \\! and &copy;  ", native.CreateReplaceEdit(escapedCharacter, "!").Apply(native.SourceMarkdown).Split('\n')[0]);
        Assert.Equal("Escaped \\* and &reg;  ", native.CreateReplaceEdit(entitySourceText, "&reg;").Apply(native.SourceMarkdown).Split('\n')[0]);
        Assert.Equal("Escaped \\* and &copy;\\", native.CreateReplaceEdit(hardBreakMarker, "\\").Apply(native.SourceMarkdown).Split('\n')[0]);
    }

    [Fact]
    public void Navigation_Helpers_Find_Native_Block_Source_Fields_By_Position_And_Name() {
        var markdown = """
# Title

> [!NOTE] Heads up
> Body

<details>
<summary>More</summary>

Inside
</details>

[^note]: Footnote

```cs
Console.WriteLine();
```

> Quote

| A | B |
| --- | --- |
| 1 | 2 |

---
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var fields = native.EnumerateBlockSourceFields().ToArray();
        var text = Assert.Single(native.EnumerateBlockSourceFields("text"));
        var paragraphTexts = native.EnumerateBlockSourceFields("paragraphText").ToArray();
        var calloutOpeningMarker = Assert.Single(native.EnumerateBlockSourceFields("calloutOpeningMarker"));
        var calloutKind = Assert.Single(native.EnumerateBlockSourceFields("calloutKind"));
        var calloutClosingMarker = Assert.Single(native.EnumerateBlockSourceFields("calloutClosingMarker"));
        var title = Assert.Single(native.EnumerateBlockSourceFields("title"));
        var calloutBody = Assert.Single(native.EnumerateBlockSourceFields("calloutBody"));
        var summary = Assert.Single(native.EnumerateBlockSourceFields("summary"));
        var detailsBody = Assert.Single(native.EnumerateBlockSourceFields("detailsBody"));
        var label = Assert.Single(native.EnumerateBlockSourceFields("label"));
        var footnoteBody = Assert.Single(native.EnumerateBlockSourceFields("footnoteBody"));
        var info = Assert.Single(native.EnumerateBlockSourceFields("infoString"));
        var content = Assert.Single(native.EnumerateBlockSourceFields("content"));
        var quoteMarker = Assert.Single(native.EnumerateBlockSourceFields("quoteMarker"));
        var quoteBody = Assert.Single(native.EnumerateBlockSourceFields("quoteBody"));
        var alignment = Assert.Single(native.EnumerateBlockSourceFields("alignmentRow"));
        var alignmentCell = native.EnumerateBlockSourceFields("alignmentCell").First();
        var thematicMarker = Assert.Single(native.EnumerateBlockSourceFields("marker"));

        Assert.Contains(fields, field => field.Name == "level" && field.Value == "1");
        Assert.Equal("Title", text.Value);
        Assert.Contains(paragraphTexts, field => field.Value == "Body" && field.SourceSpan.StartLine == 4);
        Assert.Contains(paragraphTexts, field => field.Value == "Inside" && field.SourceSpan.StartLine == 9);
        Assert.Equal("[!", calloutOpeningMarker.Value);
        Assert.Equal("note", calloutKind.Value);
        Assert.Equal("]", calloutClosingMarker.Value);
        Assert.Equal("Heads up", title.Value);
        Assert.Equal("Body", calloutBody.Value);
        Assert.Equal("More", summary.Value);
        Assert.Null(detailsBody.Value);
        Assert.Equal("note", label.Value);
        Assert.Equal("Footnote", footnoteBody.Value);
        Assert.Equal("cs", info.Value);
        Assert.Equal("Console.WriteLine();", content.Value!.Trim());
        Assert.Equal(0, quoteMarker.Index);
        Assert.Null(quoteBody.Value);
        Assert.Empty(native.EnumerateBlockSourceFields("missing"));

        AssertEquivalentField(text, native.FindBlockSourceFieldAtPosition(1, 3));
        AssertEquivalentField(calloutOpeningMarker, native.FindBlockSourceFieldAtPosition(3, 3));
        AssertEquivalentField(calloutKind, native.FindBlockSourceFieldAtPosition(3, 5));
        AssertEquivalentField(calloutClosingMarker, native.FindBlockSourceFieldAtPosition(3, 9));
        AssertEquivalentField(title, native.FindBlockSourceFieldAtPosition(3, 13));
        AssertEquivalentField(calloutBody, native.FindBlockSourceFieldAtPosition(4, 3));
        AssertEquivalentField(summary, native.FindBlockSourceFieldAtPosition(7, 10));
        AssertEquivalentField(detailsBody, native.FindBlockSourceFieldAtPosition(9, 3));
        AssertEquivalentField(label, native.FindBlockSourceFieldAtPosition(12, 4));
        AssertEquivalentField(footnoteBody, native.FindBlockSourceFieldAtPosition(12, 10));
        AssertEquivalentField(info, native.FindBlockSourceFieldAtPosition(14, 5));
        AssertEquivalentField(content, native.FindBlockSourceFieldAtPosition(15, 3));
        AssertEquivalentField(quoteMarker, native.FindBlockSourceFieldAtPosition(18, 1));
        AssertEquivalentField(quoteBody, native.FindBlockSourceFieldAtPosition(18, 3));
        AssertEquivalentField(alignmentCell, native.FindBlockSourceFieldAtPosition(21, 3));
        AssertEquivalentField(thematicMarker, native.FindBlockSourceFieldAtPosition(24, 2));
        Assert.Null(native.FindBlockSourceFieldAtPosition(2, 1));

        Assert.Contains("# New Title", native.CreateReplaceEdit(text, "New Title").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("> [!TIP] Heads up", native.CreateReplaceEdit(calloutKind, "TIP").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("> Updated body", native.CreateReplaceEdit(calloutBody, "Updated body").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("<summary>Less</summary>", native.CreateReplaceEdit(summary, "<summary>Less</summary>").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("Updated details", native.CreateReplaceEdit(detailsBody, "Updated details").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("[^note]: Updated footnote", native.CreateReplaceEdit(footnoteBody, "Updated footnote").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("```powershell", native.CreateReplaceEdit(info, "powershell").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("> Updated quote", native.CreateReplaceEdit(quoteBody, "Updated quote").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("| :---: | --- |", native.CreateReplaceEdit(alignment, "| :---: | --- |").Apply(native.SourceMarkdown), StringComparison.Ordinal);

        static void AssertEquivalentField(MarkdownNativeBlockSourceField expected, MarkdownNativeBlockSourceField? actual) {
            Assert.NotNull(actual);
            Assert.Equal(expected.Name, actual.Name);
            Assert.Equal(expected.Value, actual.Value);
            Assert.Equal(expected.SourceSpan, actual.SourceSpan);
            Assert.Same(expected.Block, actual.Block);
            Assert.Equal(expected.Index, actual.Index);
        }
    }

    [Fact]
    public void Navigation_Helpers_Find_Definition_List_Source_Fields_By_Position_And_Name() {
        var markdown = """
**Term**: Intro

  - first

Second: Other
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var definitionList = Assert.IsType<MarkdownNativeDefinitionListBlock>(Assert.Single(native.Blocks));

        var terms = native.EnumerateBlockSourceFields("definitionTerm").ToArray();
        var markers = native.EnumerateBlockSourceFields("definitionMarker").ToArray();
        var definitions = native.EnumerateBlockSourceFields("definitionBody").ToArray();
        var blankLines = native.EnumerateBlockSourceFields("definitionBlankLine").ToArray();
        var continuationIndents = native.EnumerateBlockSourceFields("definitionContinuationIndent").ToArray();

        Assert.Equal(new[] { "Term", "Second" }, terms.Select(field => field.Value).ToArray());
        Assert.Equal(new[] { 0, 1 }, terms.Select(field => field.Index).ToArray());
        Assert.Equal(new[] { ":", ":" }, markers.Select(field => field.Value).ToArray());
        Assert.Equal(new[] { 0, 1 }, markers.Select(field => field.Index).ToArray());
        Assert.Equal(new[] { "Intro\n\n- first", "Other" }, definitions.Select(field => field.Value!.Replace("\r\n", "\n")).ToArray());
        Assert.Equal(new[] { 0, 1 }, definitions.Select(field => field.Index).ToArray());
        Assert.Equal(new[] { string.Empty }, blankLines.Select(field => field.Value).ToArray());
        Assert.Equal(new[] { 0 }, blankLines.Select(field => field.Index).ToArray());
        Assert.Equal(new string?[] { null }, continuationIndents.Select(field => field.Value).ToArray());
        Assert.Equal(new[] { 0 }, continuationIndents.Select(field => field.Index).ToArray());
        Assert.All(terms, field => Assert.Same(definitionList, field.Block));
        Assert.All(markers, field => Assert.Same(definitionList, field.Block));
        Assert.All(definitions, field => Assert.Same(definitionList, field.Block));
        Assert.All(blankLines, field => Assert.Same(definitionList, field.Block));
        Assert.All(continuationIndents, field => Assert.Same(definitionList, field.Block));

        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 8), terms[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 9, 1, 9), markers[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 3, 9), definitions[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 1), blankLines[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 2), continuationIndents[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 6), terms[1].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 7, 5, 7), markers[1].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 9, 5, 13), definitions[1].SourceSpan);

        AssertEquivalentField(terms[0], native.FindBlockSourceFieldAtPosition(1, 3));
        AssertEquivalentField(markers[0], native.FindBlockSourceFieldAtPosition(1, 9));
        AssertEquivalentField(definitions[0], native.FindBlockSourceFieldAtPosition(1, 11));
        AssertEquivalentField(blankLines[0], native.FindBlockSourceFieldAtPosition(2, 1));
        AssertEquivalentField(continuationIndents[0], native.FindBlockSourceFieldAtPosition(3, 1));
        AssertEquivalentField(definitions[0], native.FindBlockSourceFieldAtPosition(3, 5));
        AssertEquivalentField(terms[1], native.FindBlockSourceFieldAtPosition(5, 2));
        AssertEquivalentField(markers[1], native.FindBlockSourceFieldAtPosition(5, 7));
        AssertEquivalentField(definitions[1], native.FindBlockSourceFieldAtPosition(5, 9));
        Assert.Null(native.FindBlockSourceFieldAtPosition(4, 1));

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Collection(
            snapshot.SourceFields,
            field => {
                Assert.Equal("definitionTerm", field.Name);
                Assert.Equal("Term", field.Value);
                Assert.Equal(0, field.Index);
            },
            field => {
                Assert.Equal("definitionMarker", field.Name);
                Assert.Equal(":", field.Value);
                Assert.Equal(0, field.Index);
            },
            field => {
                Assert.Equal("definitionBody", field.Name);
                Assert.Equal("Intro\n\n- first", field.Value!.Replace("\r\n", "\n"));
                Assert.Equal(0, field.Index);
            },
            field => {
                Assert.Equal("definitionBlankLine", field.Name);
                Assert.Equal(string.Empty, field.Value);
                Assert.Equal(0, field.Index);
            },
            field => {
                Assert.Equal("definitionContinuationIndent", field.Name);
                Assert.Null(field.Value);
                Assert.Equal(0, field.Index);
            },
            field => {
                Assert.Equal("definitionTerm", field.Name);
                Assert.Equal("Second", field.Value);
                Assert.Equal(1, field.Index);
            },
            field => {
                Assert.Equal("definitionMarker", field.Name);
                Assert.Equal(":", field.Value);
                Assert.Equal(1, field.Index);
            },
            field => {
                Assert.Equal("definitionBody", field.Name);
                Assert.Equal("Other", field.Value);
                Assert.Equal(1, field.Index);
            });

        Assert.StartsWith("**Topic**: Intro", native.CreateReplaceEdit(terms[0], "**Topic**").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("Second=> Other", native.CreateReplaceEdit(markers[1], "=>").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("Second: Updated", native.CreateReplaceEdit(definitions[1], "Updated").Apply(native.SourceMarkdown), StringComparison.Ordinal);

        static void AssertEquivalentField(MarkdownNativeBlockSourceField expected, MarkdownNativeBlockSourceField? actual) {
            Assert.NotNull(actual);
            Assert.Equal(expected.Name, actual.Name);
            Assert.Equal(expected.Value, actual.Value);
            Assert.Equal(expected.SourceSpan, actual.SourceSpan);
            Assert.Same(expected.Block, actual.Block);
            Assert.Equal(expected.Index, actual.Index);
        }
    }

    [Fact]
    public void Source_Slice_Helpers_Expose_Original_Definition_Body_Text_For_Native_Field() {
        const string markdown = "Term\r\n:   First\r\n\r\n    - item\r\nlazy continuation\r\n";
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        options.DefinitionLists = true;
        options.PreserveTrivia = true;

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));

        Assert.Equal("First\n\n- item\n  lazy continuation", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 5, 17), definitionBody.SourceSpan);

        Assert.True(native.TryCreateSourceSlice(definitionBody, out var normalizedSlice));
        Assert.Equal(MarkdownSourceTextKind.Normalized, normalizedSlice.TextKind);
        Assert.Equal("First\n\n    - item\nlazy continuation", normalizedSlice.Text);

        Assert.True(native.TryCreateOriginalSourceSlice(definitionBody, out var originalSlice));
        Assert.Equal(MarkdownSourceTextKind.Original, originalSlice.TextKind);
        Assert.Equal("First\r\n\r\n    - item\r\nlazy continuation", originalSlice.Text);

        var noTriviaOptions = MarkdownReaderOptions.CreateCommonMarkProfile();
        noTriviaOptions.DefinitionLists = true;
        var noTriviaNative = MarkdownNativeDocument.Parse(markdown, noTriviaOptions);
        var noTriviaDefinitionBody = Assert.Single(noTriviaNative.EnumerateBlockSourceFields("definitionBody"));

        Assert.True(noTriviaNative.TryCreateSourceSlice(noTriviaDefinitionBody, out _));
        Assert.False(noTriviaNative.TryCreateOriginalSourceSlice(noTriviaDefinitionBody, out _));
    }

    [Theory]
    [InlineData("**Term**: Intro\n\n  - first\n", "**Term**: Updated\n  - second\n")]
    [InlineData("Term\n:   Intro\n    - first\n", "Term\n:   Updated\n    - second\n")]
    [InlineData("Term\n:   Intro\nlazy continuation\n", "Term\n:   Updated\n    - second\n")]
    public void Definition_List_Definition_Source_Edit_Indents_Multiline_Body_For_Reparse(
        string markdown,
        string expected) {
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        options.DefinitionLists = true;
        var native = MarkdownNativeDocument.Parse(markdown, options);
        var definition = Assert.Single(native.EnumerateDefinitionListDefinitions());
        var definitionField = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));

        var edited = native.CreateReplaceEdit(definition, "Updated\n- second").Apply(native.SourceMarkdown);
        var editedViaField = native.CreateReplaceEdit(definitionField, "Updated\n- second").Apply(native.SourceMarkdown);
        var reparsed = MarkdownNativeDocument.Parse(edited, options);
        var reparsedDefinition = Assert.Single(reparsed.EnumerateDefinitionListDefinitions());

        Assert.Equal(expected, edited);
        Assert.Equal(expected, editedViaField);
        Assert.Equal("Updated\n\n- second", reparsedDefinition.Markdown.Replace("\r\n", "\n"));
        Assert.Collection(
            reparsedDefinition.Children,
            child => Assert.IsType<MarkdownNativeParagraphBlock>(child),
            child => Assert.IsType<MarkdownNativeListBlock>(child));
    }

    [Fact]
    public void Parse_Projects_Footnote_Definitions_With_Label_SourceSpan_And_Children() {
        var markdown = """
Body[^shape]

[^shape]: Footnote body

  - nested
""";

        var native = MarkdownNativeDocument.Parse(markdown);

        Assert.Equal(new[] {
            MarkdownNativeBlockKind.Paragraph,
            MarkdownNativeBlockKind.FootnoteDefinition
        }, native.Blocks.Select(block => block.Kind).ToArray());

        var footnote = Assert.IsType<MarkdownNativeFootnoteDefinitionBlock>(native.Blocks[1]);
        Assert.Equal("shape", footnote.Label);
        Assert.Equal(new MarkdownSourceSpan(3, 3, 3, 7), footnote.LabelSourceSpan);
        Assert.Equal("Footnote body\n\n- nested", footnote.Text.Replace("\r\n", "\n"));
        Assert.Collection(
            footnote.Children,
            block => Assert.Equal("Footnote body", Assert.IsType<MarkdownNativeParagraphBlock>(block).Text),
            block => Assert.Single(Assert.IsType<MarkdownNativeListBlock>(block).Items));
        Assert.Same(footnote, native.FindBlockAtPosition(3, 3));
        Assert.Same(footnote.Children[0], native.FindBlockAtPosition(3, 12));
        Assert.DoesNotContain(native.Diagnostics, diagnostic =>
            diagnostic.Id == "native.unsupported-block"
            && ReferenceEquals(diagnostic.Block, footnote));
    }

    [Fact]
    public void Parse_Projects_Footnote_Fenced_Code_SourceSpans_Remapped_Into_Native_Children() {
        var markdown = """
Lead[^note]

[^note]:
  ```ps
  Write-Host hi
  ```
""";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var footnote = Assert.IsType<MarkdownNativeFootnoteDefinitionBlock>(Assert.Single(native.Blocks, block => block.Kind == MarkdownNativeBlockKind.FootnoteDefinition));
        var code = Assert.Single(footnote.Children.OfType<MarkdownNativeCodeBlock>());

        Assert.Equal(new MarkdownSourceSpan(3, 3, 3, 6), footnote.LabelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 3, 4, 5), code.OpeningFenceSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 6, 4, 7), code.InfoStringSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 3, 5, 15), code.ContentSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(6, 3, 6, 5), code.ClosingFenceSourceSpan);
        Assert.Same(code, native.FindBlockAtPosition(5, 6));

        var snapshot = native.ToSnapshot();
        var footnoteSnapshot = Assert.Single(snapshot.Blocks, block => block.Kind == MarkdownNativeBlockKind.FootnoteDefinition);
        var codeSnapshot = Assert.Single(footnoteSnapshot.Children, block => block.Kind == MarkdownNativeBlockKind.Code);
        Assert.Equal(3, codeSnapshot.FieldSourceSpans["openingFence"]!.StartColumn);
        Assert.Equal(5, codeSnapshot.FieldSourceSpans["openingFence"]!.EndColumn);
        Assert.Equal(3, codeSnapshot.FieldSourceSpans["content"]!.StartColumn);
        Assert.Equal(15, codeSnapshot.FieldSourceSpans["content"]!.EndColumn);
        Assert.Equal(3, codeSnapshot.FieldSourceSpans["closingFence"]!.StartColumn);
        Assert.Equal(5, codeSnapshot.FieldSourceSpans["closingFence"]!.EndColumn);

        var edited = native.CreateReplaceEdit(code.ContentSourceSpan!.Value, "Write-Output hi").Apply(native.SourceMarkdown);
        Assert.Contains("  Write-Output hi", edited, StringComparison.Ordinal);
        Assert.DoesNotContain("Write-Host hi", edited, StringComparison.Ordinal);
    }

    [Fact]
    public void Parse_Projects_Definition_Lists_With_Terms_Definitions_And_Children() {
        var markdown = """
**Term**: Intro

  - first
  - second
""";

        var native = MarkdownNativeDocument.Parse(markdown);

        var definitionList = Assert.IsType<MarkdownNativeDefinitionListBlock>(Assert.Single(native.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var term = Assert.Single(group.Terms);
        var definition = Assert.Single(group.Definitions);

        Assert.Equal(MarkdownNativeBlockKind.DefinitionList, definitionList.Kind);
        Assert.Equal("Term", term.Text);
        Assert.Equal("**Term**", term.Markdown);
        Assert.Equal("**Term**", term.TermObject.Markdown);
        Assert.Same(term.TermObject.Inlines, term.Term);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 8), term.SourceSpan);
        Assert.Equal(term.SourceSpan, term.TermObject.SourceSpan);
        Assert.Contains(term.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Strong && inline.Text == "Term");
        Assert.Equal("Intro\n\n- first\n- second", definition.Markdown.Replace("\r\n", "\n"));
        Assert.Collection(
            definition.Children,
            block => Assert.Equal("Intro", Assert.IsType<MarkdownNativeParagraphBlock>(block).Text),
            block => Assert.Equal(2, Assert.IsType<MarkdownNativeListBlock>(block).Items.Count));
        Assert.Equal(definition.Children, definitionList.Children);
        Assert.Same(definition.Children[1], native.FindBlockAtPosition(3, 3));
        Assert.Contains(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Strong && inline.Text == "Term");
        Assert.DoesNotContain(native.Diagnostics, diagnostic =>
            diagnostic.Id == "native.unsupported-block"
            && ReferenceEquals(diagnostic.Block, definitionList));
    }

    [Fact]
    public void Source_Edit_Helpers_Replace_Definition_List_Parts_And_Table_Cells() {
        var markdown = """
**Term**: Intro

  - first
  - second

| Name | Value |
| --- | --- |
| CPU | 42 |
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var definitionList = Assert.IsType<MarkdownNativeDefinitionListBlock>(native.Blocks[0]);
        var group = Assert.Single(definitionList.Groups);
        var term = Assert.Single(group.Terms);
        var definition = Assert.Single(group.Definitions);
        var table = Assert.IsType<MarkdownNativeTableBlock>(native.Blocks[1]);
        var valueCell = table.Rows[0][1];

        Assert.Equal("**Topic**: Intro", native.CreateReplaceEdit(term, "**Topic**").Apply(native.SourceMarkdown).Split('\n')[0]);
        Assert.Equal("Updated", native.CreateReplaceEdit(definition, "Updated").Apply(native.SourceMarkdown).Split('\n')[0].Substring("**Term**: ".Length));
        Assert.StartsWith("New: Body", native.CreateReplaceEdit(group, "New: Body").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Equal("| CPU | 84 |", native.CreateReplaceEdit(valueCell, "84").Apply(native.SourceMarkdown).Split('\n')[7]);
    }

    [Fact]
    public void Source_Edit_Helpers_Roundtrip_Definition_List_And_Table_Cell_Edits_Against_Original_Source() {
        const string markdown = "**Term**: Intro\r\n\r\n  - first\r\n  - second\r\n\r\n| Name | Value |\r\n| --- | --- |\r\n| CPU | 42 |\r\n";
        var native = MarkdownNativeDocument.Parse(
            markdown,
            new MarkdownReaderOptions { PreserveTrivia = true });
        var definitionList = Assert.IsType<MarkdownNativeDefinitionListBlock>(native.Blocks[0]);
        var group = Assert.Single(definitionList.Groups);
        var term = Assert.Single(group.Terms);
        var definition = Assert.Single(group.Definitions);
        var table = Assert.IsType<MarkdownNativeTableBlock>(native.Blocks[1]);
        var valueCell = table.Rows[0][1];

        var roundtrip = native.WriteWithSourceEdits(new[] {
            native.CreateReplaceEdit(term, "**Topic**"),
            native.CreateReplaceEdit(definition, "Updated\r\n\r\n  - keep"),
            native.CreateReplaceEdit(valueCell, "84")
        });

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("**Topic**: Updated\r\n\r\n  - keep\r\n\r\n| Name | Value |\r\n| --- | --- |\r\n| CPU | 84 |\r\n", roundtrip.Markdown);
    }

    [Fact]
    public void ToSnapshot_Projects_Definition_List_Groups_Terms_And_Definition_Children() {
        var native = MarkdownNativeDocument.Parse("""
Term: Intro

  - first
""");

        var snapshot = native.ToSnapshot();
        var definitionList = Assert.Single(snapshot.Blocks);
        var group = Assert.Single(definitionList.DefinitionGroups);
        var term = Assert.Single(group.Terms);
        var definition = Assert.Single(group.Definitions);

        Assert.Equal(MarkdownNativeBlockKind.DefinitionList, definitionList.Kind);
        Assert.Equal("Term", term.Text);
        Assert.Equal("Term", term.Markdown);
        Assert.Equal("Intro\n\n- first", definition.Markdown.Replace("\r\n", "\n"));
        Assert.Collection(
            definition.Children,
            block => Assert.Equal(MarkdownNativeBlockKind.Paragraph, block.Kind),
            block => Assert.Equal(MarkdownNativeBlockKind.List, block.Kind));
        Assert.Equal(definition.Children.Select(block => block.Id).ToArray(), definitionList.Children.Select(block => block.Id).ToArray());
    }

    [Fact]
    public void Parse_Projects_Thematic_Break_As_FirstClass_Native_Block() {
        var native = MarkdownNativeDocument.Parse("""
Before

***

After
""");

        Assert.Equal(new[] {
            MarkdownNativeBlockKind.Paragraph,
            MarkdownNativeBlockKind.ThematicBreak,
            MarkdownNativeBlockKind.Paragraph
        }, native.Blocks.Select(block => block.Kind).ToArray());

        var thematicBreak = Assert.IsType<MarkdownNativeThematicBreakBlock>(native.Blocks[1]);
        Assert.Equal("---", thematicBreak.Marker);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 3), thematicBreak.SourceSpan);
        Assert.Same(thematicBreak, native.FindBlockAtPosition(3, 2));
        Assert.DoesNotContain(native.Diagnostics, diagnostic =>
            diagnostic.Id == "native.unsupported-block"
            && ReferenceEquals(diagnostic.Block, thematicBreak));
    }

    [Fact]
    public void Parse_Exposes_Reference_Link_Definitions_On_Native_Document_And_Snapshot() {
        var native = MarkdownNativeDocument.Parse("""
[hero]: https://example.com/docs "Docs title"

[hero]
""");

        var definition = Assert.Single(native.ReferenceLinkDefinitions);
        Assert.Equal("hero", definition.Label);
        Assert.Equal("https://example.com/docs", definition.Url);
        Assert.Equal("Docs title", definition.Title);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 45), definition.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 2, 1, 5), definition.LabelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 1), definition.OpeningMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 7), definition.SeparatorMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 9, 1, 32), definition.UrlSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 35, 1, 44), definition.TitleSourceSpan);

        var snapshotDefinition = Assert.Single(native.ToSnapshot().ReferenceLinkDefinitions);
        Assert.Equal("hero", snapshotDefinition.Label);
        Assert.Equal("https://example.com/docs", snapshotDefinition.Url);
        Assert.Equal("Docs title", snapshotDefinition.Title);
        Assert.Equal(1, snapshotDefinition.SourceSpan!.StartLine);
        Assert.Equal(1, snapshotDefinition.SourceSpan.StartColumn);
        Assert.Equal(45, snapshotDefinition.SourceSpan.EndColumn);
        Assert.Equal(1, snapshotDefinition.LabelSourceSpan!.StartLine);
        Assert.Equal(2, snapshotDefinition.LabelSourceSpan.StartColumn);
        Assert.Equal(5, snapshotDefinition.LabelSourceSpan.EndColumn);
        Assert.Equal(1, snapshotDefinition.OpeningMarkerSourceSpan!.StartLine);
        Assert.Equal(1, snapshotDefinition.OpeningMarkerSourceSpan.StartColumn);
        Assert.Equal(1, snapshotDefinition.OpeningMarkerSourceSpan.EndColumn);
        Assert.Equal(6, snapshotDefinition.SeparatorMarkerSourceSpan!.StartColumn);
        Assert.Equal(7, snapshotDefinition.SeparatorMarkerSourceSpan.EndColumn);
        Assert.Equal(9, snapshotDefinition.UrlSourceSpan!.StartColumn);
        Assert.Equal(32, snapshotDefinition.UrlSourceSpan.EndColumn);
        Assert.Equal(35, snapshotDefinition.TitleSourceSpan!.StartColumn);
        Assert.Equal(44, snapshotDefinition.TitleSourceSpan.EndColumn);
        Assert.Equal(
            new[] { "openingMarker", "label", "separatorMarker", "url", "title" },
            snapshotDefinition.SourceFields.Select(field => field.Name).ToArray());
        Assert.Equal("[", snapshotDefinition.SourceFields[0].Value);
        Assert.Equal("hero", snapshotDefinition.SourceFields[1].Value);
        Assert.Equal("]:", snapshotDefinition.SourceFields[2].Value);
        Assert.Equal("https://example.com/docs", snapshotDefinition.SourceFields[3].Value);
        Assert.Equal("Docs title", snapshotDefinition.SourceFields[4].Value);
    }

    [Fact]
    public void Parse_Exposes_Multiline_Reference_Definition_Label_Frame_SourceSpans() {
        var native = MarkdownNativeDocument.Parse("""
[Foo
  bar]: https://example.com

[foo bar]
""");

        var definition = Assert.Single(native.ReferenceLinkDefinitions);
        Assert.Equal("foo bar", definition.Label);
        Assert.Equal(new MarkdownSourceSpan(1, 2, 2, 5), definition.LabelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 1), definition.OpeningMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 6, 2, 7), definition.SeparatorMarkerSourceSpan);

        var snapshotDefinition = Assert.Single(native.ToSnapshot().ReferenceLinkDefinitions);
        Assert.Equal(1, snapshotDefinition.OpeningMarkerSourceSpan!.StartLine);
        Assert.Equal(1, snapshotDefinition.OpeningMarkerSourceSpan.StartColumn);
        Assert.Equal(2, snapshotDefinition.SeparatorMarkerSourceSpan!.StartLine);
        Assert.Equal(6, snapshotDefinition.SeparatorMarkerSourceSpan.StartColumn);
        Assert.Equal(7, snapshotDefinition.SeparatorMarkerSourceSpan.EndColumn);
    }

    [Fact]
    public void Parse_Exposes_Multiline_Reference_Definition_Label_With_Line_Leading_Separator() {
        var native = MarkdownNativeDocument.Parse("""
[
foo
]: /url
bar
""", MarkdownReaderOptions.CreateCommonMarkProfile());

        var definition = Assert.Single(native.ReferenceLinkDefinitions);
        Assert.Equal("foo", definition.Label);
        Assert.Equal("/url", definition.Url);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 3), definition.LabelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 1), definition.OpeningMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 2), definition.SeparatorMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 4, 3, 7), definition.UrlSourceSpan);

        var fields = native.EnumerateReferenceLinkDefinitionFields().ToArray();
        Assert.Equal(
            new[] { "openingMarker", "label", "separatorMarker", "url" },
            fields.Select(field => field.Name).ToArray());
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 3), fields[1].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 2), fields[2].SourceSpan);
        Assert.Equal("separatorMarker", native.FindReferenceLinkDefinitionFieldAtPosition(3, 1)!.Name);
        Assert.Equal("url", native.FindReferenceLinkDefinitionFieldAtPosition(3, 5)!.Name);
    }

    [Fact]
    public void Parse_Keeps_Paragraph_Interrupting_Reference_Definition_Text_Literal() {
        var native = MarkdownNativeDocument.Parse("""
Foo
[bar]: /baz

[bar]
""", MarkdownReaderOptions.CreateCommonMarkProfile());

        Assert.Empty(native.ReferenceLinkDefinitions);
        Assert.Equal(
            new[] { MarkdownNativeBlockKind.Paragraph, MarkdownNativeBlockKind.Paragraph },
            native.Blocks.Select(block => block.Kind).ToArray());
        Assert.Null(native.FindReferenceLinkDefinitionAtPosition(2, 1));
        Assert.Null(native.FindReferenceLinkDefinitionFieldAtPosition(2, 2));
        Assert.Same(native.Blocks[0], native.FindBlockAtPosition(2, 2));
    }

    [Fact]
    public void Parse_Resolves_BlockQuote_Reference_Definitions_For_Earlier_Paragraphs() {
        var native = MarkdownNativeDocument.Parse("""
[foo]

> [foo]: /url
""", MarkdownReaderOptions.CreateCommonMarkProfile());

        var definition = Assert.Single(native.ReferenceLinkDefinitions);
        Assert.Equal("foo", definition.Label);
        Assert.Equal("/url", definition.Url);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 13), definition.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 3, 3, 3), definition.OpeningMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 4, 3, 6), definition.LabelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 8), definition.SeparatorMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 10, 3, 13), definition.UrlSourceSpan);

        Assert.Equal(
            new[] { MarkdownNativeBlockKind.Paragraph, MarkdownNativeBlockKind.Quote },
            native.Blocks.Select(block => block.Kind).ToArray());
        Assert.Same(native.Blocks[1], native.FindBlockAtPosition(3, 1));
        Assert.Same(definition, native.FindReferenceLinkDefinitionAtPosition(3, 4));
        Assert.Equal("url", native.FindReferenceLinkDefinitionFieldAtPosition(3, 10)!.Name);
        Assert.Equal(
            CommonMarkHtmlComparison.Normalize("<p><a href=\"/url\">foo</a></p>\n<blockquote>\n</blockquote>"),
            CommonMarkHtmlComparison.Normalize(native.Document.ToHtmlFragment(CommonMarkHtmlComparison.CreatePlainHtmlOptions())));
    }

    [Fact]
    public void Native_Document_Enumerates_Reference_Definition_SourceFields_And_Position_Lookup() {
        var native = MarkdownNativeDocument.Parse("""
[hero]: https://example.com/docs "Docs title"

[hero]
""");

        var definition = Assert.Single(native.ReferenceLinkDefinitions);
        var fields = native.EnumerateReferenceLinkDefinitionFields().ToArray();

        Assert.Equal(
            new[] { "openingMarker", "label", "separatorMarker", "url", "title" },
            fields.Select(field => field.Name).ToArray());
        Assert.All(fields, field => Assert.Same(definition, field.Definition));
        Assert.Equal("[", fields[0].Value);
        Assert.Equal("hero", fields[1].Value);
        Assert.Equal("]:", fields[2].Value);
        Assert.Equal("https://example.com/docs", fields[3].Value);
        Assert.Equal("Docs title", fields[4].Value);

        var label = Assert.Single(native.EnumerateReferenceLinkDefinitionFields("label"));
        Assert.Equal(new MarkdownSourceSpan(1, 2, 1, 5), label.SourceSpan);
        Assert.Empty(native.EnumerateReferenceLinkDefinitionFields("missing"));
        Assert.Same(definition, native.FindReferenceLinkDefinitionAtPosition(1, 1));
        Assert.Same(definition, native.FindReferenceLinkDefinitionAtPosition(1, 40));
        Assert.Equal("openingMarker", native.FindReferenceLinkDefinitionFieldAtPosition(1, 1)!.Name);
        Assert.Equal("label", native.FindReferenceLinkDefinitionFieldAtPosition(1, 3)!.Name);
        Assert.Equal("separatorMarker", native.FindReferenceLinkDefinitionFieldAtPosition(1, 6)!.Name);
        Assert.Equal("url", native.FindReferenceLinkDefinitionFieldAtPosition(1, 12)!.Name);
        Assert.Equal("title", native.FindReferenceLinkDefinitionFieldAtPosition(1, 39)!.Name);
        Assert.Null(native.FindReferenceLinkDefinitionAtPosition(3, 1));
        Assert.Null(native.FindReferenceLinkDefinitionFieldAtPosition(3, 1));
    }

    [Fact]
    public void Source_Edit_Helpers_Replace_Reference_Definition_And_Preserve_Surrounding_Source() {
        var markdown = """
Before [hero]

[hero]: https://example.com/docs "Docs title"

After
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var definition = Assert.Single(native.ReferenceLinkDefinitions);

        var edit = native.CreateReplaceEdit(definition, "[hero]: https://example.com/new \"New title\"");
        var updated = edit.Apply(native.SourceMarkdown);

        Assert.Equal(NormalizeLineEndings("""
Before [hero]

[hero]: https://example.com/new "New title"

After
"""), NormalizeLineEndings(updated));

        var openingMarker = Assert.Single(native.EnumerateReferenceLinkDefinitionFields("openingMarker"));
        var openingMarkerEdit = native.CreateReplaceEdit(openingMarker, "[ref-");
        Assert.Contains("[ref-hero]: https://example.com/docs", openingMarkerEdit.Apply(native.SourceMarkdown), StringComparison.Ordinal);

        var separatorMarker = Assert.Single(native.EnumerateReferenceLinkDefinitionFields("separatorMarker"));
        var separatorMarkerEdit = native.CreateReplaceEdit(separatorMarker, "]: ");
        Assert.Contains("[hero]:  https://example.com/docs", separatorMarkerEdit.Apply(native.SourceMarkdown), StringComparison.Ordinal);
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
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 3), code.OpeningFenceSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 4, 3, 9), code.InfoStringSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 21), code.ContentSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 3), code.ClosingFenceSourceSpan);

        var table = Assert.IsType<MarkdownNativeTableBlock>(native.Blocks[2]);
        Assert.Equal("Name", table.HeaderCells[0].Text);
        Assert.Equal("42", table.Rows[0][1].Text);
        Assert.Equal(7, table.SourceSpan!.Value.StartLine);

        var visual = Assert.IsType<MarkdownNativeVisualBlock>(native.Blocks[3]);
        Assert.Equal(MarkdownSemanticKinds.Chart, visual.SemanticKind);
        Assert.Equal("ix-chart", visual.Language);
        Assert.Equal("{\"type\":\"bar\"}", visual.Content);
        Assert.Equal(11, visual.SourceSpan!.Value.StartLine);
        Assert.Equal(new MarkdownSourceSpan(11, 1, 11, 3), visual.OpeningFenceSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(11, 4, 11, 11), visual.InfoStringSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(12, 1, 12, 14), visual.ContentSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(13, 1, 13, 3), visual.ClosingFenceSourceSpan);
        var snapshot = native.ToSnapshot();
        Assert.Equal(1, snapshot.Blocks[1].FieldSourceSpans["openingFence"]!.StartColumn);
        Assert.Equal(3, snapshot.Blocks[1].FieldSourceSpans["openingFence"]!.EndColumn);
        Assert.Equal(4, snapshot.Blocks[1].FieldSourceSpans["infoString"]!.StartColumn);
        Assert.Equal(9, snapshot.Blocks[1].FieldSourceSpans["infoString"]!.EndColumn);
        Assert.Equal(1, snapshot.Blocks[1].FieldSourceSpans["closingFence"]!.StartColumn);
        Assert.Equal(3, snapshot.Blocks[1].FieldSourceSpans["closingFence"]!.EndColumn);
        Assert.Equal(1, snapshot.Blocks[3].FieldSourceSpans["openingFence"]!.StartColumn);
        Assert.Equal(3, snapshot.Blocks[3].FieldSourceSpans["openingFence"]!.EndColumn);
        Assert.Equal(4, snapshot.Blocks[3].FieldSourceSpans["infoString"]!.StartColumn);
        Assert.Equal(11, snapshot.Blocks[3].FieldSourceSpans["infoString"]!.EndColumn);
        Assert.Equal(1, snapshot.Blocks[3].FieldSourceSpans["closingFence"]!.StartColumn);
        Assert.Equal(3, snapshot.Blocks[3].FieldSourceSpans["closingFence"]!.EndColumn);

        var reticked = native.CreateReplaceEdit(code.OpeningFenceSourceSpan!.Value, "````").Apply(native.SourceMarkdown);
        Assert.Contains("````csharp", reticked, StringComparison.Ordinal);
        var visualClose = native.CreateReplaceEdit(visual.ClosingFenceSourceSpan!.Value, "````").Apply(native.SourceMarkdown);
        Assert.Contains("{\"type\":\"bar\"}\n````", visualClose, StringComparison.Ordinal);
        Assert.Same(visual, native.FindBlockAtLine(12));
    }

    [Fact]
    public void Parse_Projects_Unclosed_Fenced_Code_Exposes_Opening_But_Not_Closing_Fence_SourceSpan() {
        var native = MarkdownNativeDocument.Parse("```text\nbody");
        var code = Assert.IsType<MarkdownNativeCodeBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 3), code.OpeningFenceSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 4, 1, 7), code.InfoStringSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 4), code.ContentSourceSpan);
        Assert.Null(code.ClosingFenceSourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(1, snapshot.FieldSourceSpans["openingFence"]!.StartColumn);
        Assert.Null(snapshot.FieldSourceSpans["closingFence"]);
    }

    [Fact]
    public void Parse_Projects_Nested_Fenced_Code_Marker_SourceSpans_Remapped_Through_Quote_And_List() {
        var markdown = """
> ```csharp
> Console.WriteLine(1);
> ```

- item
  ```json
  {"ok":true}
  ```
""";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        var quote = Assert.IsType<MarkdownNativeQuoteBlock>(native.Blocks[0]);
        var quotedCode = Assert.IsType<MarkdownNativeCodeBlock>(Assert.Single(quote.Children));
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 5), quotedCode.OpeningFenceSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 11), quotedCode.InfoStringSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 23), quotedCode.ContentSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 3, 3, 5), quotedCode.ClosingFenceSourceSpan);

        var list = Assert.IsType<MarkdownNativeListBlock>(native.Blocks[1]);
        var listCode = Assert.Single(Assert.Single(list.Items).Children.OfType<MarkdownNativeCodeBlock>());
        Assert.Equal(new MarkdownSourceSpan(6, 3, 6, 5), listCode.OpeningFenceSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(6, 6, 6, 9), listCode.InfoStringSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(7, 3, 7, 13), listCode.ContentSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(8, 3, 8, 5), listCode.ClosingFenceSourceSpan);

        var snapshot = native.ToSnapshot();
        var quotedSnapshot = Assert.Single(snapshot.Blocks[0].Children);
        Assert.Equal(3, quotedSnapshot.FieldSourceSpans["openingFence"]!.StartColumn);
        Assert.Equal(5, quotedSnapshot.FieldSourceSpans["openingFence"]!.EndColumn);
        Assert.Equal(3, quotedSnapshot.FieldSourceSpans["closingFence"]!.StartColumn);
        Assert.Equal(5, quotedSnapshot.FieldSourceSpans["closingFence"]!.EndColumn);

        var listSnapshot = Assert.Single(Assert.Single(snapshot.Blocks[1].Items).Children, child => child.Kind == MarkdownNativeBlockKind.Code);
        Assert.Equal(3, listSnapshot.FieldSourceSpans["openingFence"]!.StartColumn);
        Assert.Equal(5, listSnapshot.FieldSourceSpans["openingFence"]!.EndColumn);
        Assert.Equal(3, listSnapshot.FieldSourceSpans["closingFence"]!.StartColumn);
        Assert.Equal(5, listSnapshot.FieldSourceSpans["closingFence"]!.EndColumn);

        var quotedReticked = native.CreateReplaceEdit(quotedCode.OpeningFenceSourceSpan!.Value, "````").Apply(native.SourceMarkdown);
        Assert.Contains("> ````csharp", quotedReticked, StringComparison.Ordinal);

        var listReticked = native.CreateReplaceEdit(listCode.ClosingFenceSourceSpan!.Value, "````").Apply(native.SourceMarkdown);
        Assert.Contains("\n  ````", listReticked, StringComparison.Ordinal);
    }

    [Fact]
    public void Parse_Projects_Html_Blocks_With_SourceSpans_Snapshots_And_SourceEdits() {
        var markdown = """
Before

<section data-kind="note">Raw</section>

<!-- keep
this comment
-->

After
""";

        var native = MarkdownNativeDocument.Parse(markdown);

        Assert.Equal(new[] {
            MarkdownNativeBlockKind.Paragraph,
            MarkdownNativeBlockKind.Html,
            MarkdownNativeBlockKind.Html,
            MarkdownNativeBlockKind.Paragraph
        }, native.Blocks.Select(block => block.Kind).ToArray());

        var html = Assert.IsType<MarkdownNativeHtmlBlock>(native.Blocks[1]);
        Assert.False(html.IsComment);
        Assert.Equal("<section data-kind=\"note\">Raw</section>", html.Html);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 39), html.SourceSpan);
        Assert.Same(html, native.FindBlockAtPosition(3, 10));

        var comment = Assert.IsType<MarkdownNativeHtmlBlock>(native.Blocks[2]);
        Assert.True(comment.IsComment);
        Assert.Equal("<!-- keep\nthis comment\n-->", comment.Html);
        Assert.Equal(" keep\nthis comment", comment.CommentBody);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 7, 3), comment.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 4), comment.OpeningMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 5, 6, 12), comment.BodySourceSpan);
        Assert.Equal(new MarkdownSourceSpan(7, 1, 7, 3), comment.ClosingMarkerSourceSpan);
        Assert.Same(comment, native.FindBlockAtLine(6));

        var snapshot = native.ToSnapshot();
        Assert.Equal("<section data-kind=\"note\">Raw</section>", snapshot.Blocks[1].Markdown);
        Assert.Equal("false", snapshot.Blocks[1].Fields["isComment"]);
        Assert.Equal("true", snapshot.Blocks[2].Fields["isComment"]);
        Assert.Equal(5, snapshot.Blocks[2].SourceSpan!.StartLine);
        Assert.Equal(7, snapshot.Blocks[2].SourceSpan.EndLine);
        Assert.Equal(5, snapshot.Blocks[2].FieldSourceSpans["htmlCommentOpeningMarker"]!.StartLine);
        Assert.Equal(4, snapshot.Blocks[2].FieldSourceSpans["htmlCommentOpeningMarker"]!.EndColumn);
        Assert.Equal(12, snapshot.Blocks[2].FieldSourceSpans["htmlCommentBody"]!.EndColumn);
        Assert.Equal(7, snapshot.Blocks[2].FieldSourceSpans["htmlCommentClosingMarker"]!.StartLine);

        var edit = native.CreateReplaceEdit(comment, "<!-- updated -->");
        var updated = edit.Apply(native.SourceMarkdown);

        Assert.Equal(NormalizeLineEndings("""
Before

<section data-kind="note">Raw</section>

<!-- updated -->

After
"""), updated);
        Assert.DoesNotContain("this comment", updated);
        Assert.DoesNotContain(native.Diagnostics, diagnostic =>
            diagnostic.Id == "native.unsupported-block"
            && ReferenceEquals(diagnostic.Block, html));
        Assert.DoesNotContain(native.Diagnostics, diagnostic =>
            diagnostic.Id == "native.unsupported-block"
            && ReferenceEquals(diagnostic.Block, comment));
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
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 15), table.AlignmentRowSourceSpan);
        Assert.Equal(2, table.AlignmentCells.Count);
        Assert.Equal(":---", table.AlignmentCells[0].Markdown);
        Assert.Equal("---:", table.AlignmentCells[1].Markdown);
        Assert.Equal(ColumnAlignment.Left, table.AlignmentCells[0].Alignment);
        Assert.Equal(ColumnAlignment.Right, table.AlignmentCells[1].Alignment);
        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 6), table.AlignmentCells[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 10, 2, 13), table.AlignmentCells[1].SourceSpan);
        Assert.Equal(9, table.Pipes.Count);
        Assert.Equal(-1, table.Pipes[0].RowIndex);
        Assert.Equal(0, table.Pipes[0].ColumnIndex);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 1), table.Pipes[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 8, 1, 8), table.Pipes[1].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 16, 1, 16), table.Pipes[2].SourceSpan);
        Assert.Equal(-2, table.Pipes[3].RowIndex);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 1), table.Pipes[3].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 8, 2, 8), table.Pipes[4].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 15, 2, 15), table.Pipes[5].SourceSpan);
        Assert.Equal(0, table.Pipes[6].RowIndex);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 1), table.Pipes[6].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 7), table.Pipes[7].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 12, 3, 12), table.Pipes[8].SourceSpan);

        var snapshot = native.ToSnapshot().Blocks[0];
        Assert.Equal(2, snapshot.FieldSourceSpans["alignmentRow"]!.StartLine);
        Assert.Equal(1, snapshot.FieldSourceSpans["alignmentRow"]!.StartColumn);
        Assert.Equal(15, snapshot.FieldSourceSpans["alignmentRow"]!.EndColumn);
        var alignmentCells = table.EnumerateSourceFields("alignmentCell").ToArray();
        Assert.Equal(2, alignmentCells.Length);
        Assert.Equal(0, alignmentCells[0].Index);
        Assert.Equal(1, alignmentCells[1].Index);
        Assert.Equal(":---", alignmentCells[0].Value);
        Assert.Equal("---:", alignmentCells[1].Value);
        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 6), alignmentCells[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 10, 2, 13), alignmentCells[1].SourceSpan);
        var tablePipes = table.EnumerateSourceFields("tablePipe").ToArray();
        Assert.Equal(9, tablePipes.Length);
        Assert.Equal(0, tablePipes[0].Index);
        Assert.Equal(8, tablePipes[8].Index);
        Assert.All(tablePipes, pipe => Assert.Equal("|", pipe.Value));
        Assert.Equal(new MarkdownSourceSpan(1, 8, 1, 8), tablePipes[1].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 8, 2, 8), tablePipes[4].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 7), tablePipes[7].SourceSpan);
        var selectedAlignmentCell = native.FindBlockSourceFieldAtPosition(2, 3);
        Assert.NotNull(selectedAlignmentCell);
        Assert.Equal("alignmentCell", selectedAlignmentCell!.Name);
        Assert.Equal(alignmentCells[0].SourceSpan, selectedAlignmentCell.SourceSpan);
        var selectedPipe = native.FindBlockSourceFieldAtPosition(1, 8);
        Assert.NotNull(selectedPipe);
        Assert.Equal("tablePipe", selectedPipe!.Name);
        Assert.Equal(tablePipes[1].SourceSpan, selectedPipe.SourceSpan);

        var snapshotAlignmentCells = snapshot.EnumerateSourceFields("alignmentCell").ToArray();
        Assert.Equal(2, snapshotAlignmentCells.Length);
        Assert.Equal(0, snapshotAlignmentCells[0].Index);
        Assert.Equal(1, snapshotAlignmentCells[1].Index);
        Assert.Equal(3, snapshotAlignmentCells[0].SourceSpan.StartColumn);
        Assert.Equal(13, snapshotAlignmentCells[1].SourceSpan.EndColumn);
        var snapshotPipes = snapshot.EnumerateSourceFields("tablePipe").ToArray();
        Assert.Equal(9, snapshotPipes.Length);
        Assert.Equal(0, snapshotPipes[0].Index);
        Assert.Equal(8, snapshotPipes[8].Index);
        Assert.Equal(8, snapshotPipes[1].SourceSpan.StartColumn);
        Assert.Equal(12, snapshotPipes[8].SourceSpan.EndColumn);

        var edited = native.CreateReplaceEdit(table.AlignmentRowSourceSpan!.Value, "| :---: | --- |").Apply(native.SourceMarkdown);
        Assert.Equal("| :---: | --- |", edited.Split('\n')[1]);
        var editedCell = native.CreateReplaceEdit(alignmentCells[0], ":---:").Apply(native.SourceMarkdown);
        Assert.Equal("| :---: | ---: |", editedCell.Split('\n')[1]);
        var editedPipe = native.CreateReplaceEdit(tablePipes[1], "||").Apply(native.SourceMarkdown);
        Assert.Equal("| Name || Value |", editedPipe.Split('\n')[0]);
    }

    [Fact]
    public void Parse_Table_Pipe_Source_Fields_Ignore_Escaped_And_CodeSpan_Pipes() {
        var markdown = """
| Name | Value |
| --- | --- |
| CPU \| spare | `a|b` |
""";

        var native = MarkdownNativeDocument.Parse(markdown);

        var table = Assert.IsType<MarkdownNativeTableBlock>(Assert.Single(native.Blocks));
        var tablePipes = table.EnumerateSourceFields("tablePipe").ToArray();
        Assert.Equal(9, tablePipes.Length);
        Assert.Equal(new[] { 1, 8, 16 }, tablePipes.Where(pipe => pipe.SourceSpan.StartLine == 1).Select(pipe => pipe.SourceSpan.StartColumn!.Value).ToArray());
        Assert.Equal(new[] { 1, 7, 13 }, tablePipes.Where(pipe => pipe.SourceSpan.StartLine == 2).Select(pipe => pipe.SourceSpan.StartColumn!.Value).ToArray());
        Assert.Equal(new[] { 1, 16, 24 }, tablePipes.Where(pipe => pipe.SourceSpan.StartLine == 3).Select(pipe => pipe.SourceSpan.StartColumn!.Value).ToArray());
        Assert.NotEqual("tablePipe", native.FindBlockSourceFieldAtPosition(3, 8)?.Name);
        Assert.NotEqual("tablePipe", native.FindBlockSourceFieldAtPosition(3, 20)?.Name);
    }

    [Fact]
    public void Parse_Projects_Empty_Table_Cell_SourceSpans_Into_Native_Snapshots_And_Edits() {
        var markdown = """
| Name |  |
| --- | --- |
| One |  |
""";

        var native = MarkdownNativeDocument.Parse(markdown);

        var table = Assert.IsType<MarkdownNativeTableBlock>(Assert.Single(native.Blocks));
        var headerEmpty = table.HeaderCells[1];
        var bodyEmpty = table.Rows[0][1];

        Assert.Equal(string.Empty, headerEmpty.Text);
        Assert.Equal(string.Empty, bodyEmpty.Text);
        Assert.Equal(new MarkdownSourceSpan(1, 9, 1, 10), headerEmpty.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 8, 3, 9), bodyEmpty.SourceSpan);

        var snapshot = native.ToSnapshot().Blocks[0];
        Assert.Equal(9, snapshot.HeaderCells[1].SourceSpan!.StartColumn);
        Assert.Equal(10, snapshot.HeaderCells[1].SourceSpan!.EndColumn);
        Assert.Equal(8, snapshot.Rows[0][1].SourceSpan!.StartColumn);
        Assert.Equal(9, snapshot.Rows[0][1].SourceSpan!.EndColumn);

        var edited = native.CreateReplaceEdit(bodyEmpty.SourceSpan!.Value, " 2 ").Apply(native.SourceMarkdown);
        Assert.Equal("| One | 2 |", edited.Split('\n')[2]);
    }

    [Fact]
    public void Parse_Projects_Structured_Table_Cell_Blocks_As_Native_Children() {
        var markdown = """
| Name | Detail |
| --- | --- |
| CPU | - high<br>- low |
""";

        var native = MarkdownNativeDocument.Parse(markdown);

        var table = Assert.IsType<MarkdownNativeTableBlock>(Assert.Single(native.Blocks));
        var detailCell = table.Rows[0][1];
        Assert.Empty(detailCell.InlineRuns);
        var list = Assert.IsType<MarkdownNativeListBlock>(Assert.Single(detailCell.Children));
        Assert.Equal(new[] { "high", "low" }, list.Items.Select(item => item.Text).ToArray());
        Assert.Same(list, native.FindBlockById(list.Id));
        Assert.Equal(new[] { table.Id, list.Id }, native.GetBlockPath(list.Id).Select(block => block.Id).ToArray());
        Assert.Contains(native.DescendantBlocksAndSelf(), block => ReferenceEquals(block, list));

        var snapshotCell = native.ToSnapshot().Blocks[0].Rows[0][1];
        Assert.Empty(snapshotCell.Inlines);
        var snapshotList = Assert.Single(snapshotCell.Children);
        Assert.Equal(MarkdownNativeBlockKind.List, snapshotList.Kind);
        Assert.Equal(new[] { "high", "low" }, snapshotList.Items.Select(item => item.Text).ToArray());
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
    public void Parse_Projects_Native_Inlines_With_SourceSpans_And_Metadata() {
        var markdown = """
# Native **AST** [docs](https://example.com "Docs")

Paragraph with `code`, footnote[^note], and ![Alt](img.png "Img").

[^note]: Footnote body
""";

        var native = MarkdownNativeDocument.Parse(markdown);

        var heading = Assert.IsType<MarkdownNativeHeadingBlock>(native.Blocks[0]);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 1), heading.LevelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 51), heading.TextSourceSpan);
        Assert.Contains(heading.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Strong && inline.Text == "AST");

        var link = Assert.Single(heading.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Link);
        Assert.Equal("docs", link.Text);
        Assert.Equal("https://example.com", link.GetMetadata("target"));
        Assert.Equal("Docs", link.GetMetadata("title"));
        Assert.True(link.SourceSpan.HasValue);
        Assert.Same(link, native.FindInlineById(link.Id));
        Assert.Same(link, native.FindInlineAtPosition(link.SourceSpan.Value.StartLine, link.SourceSpan.Value.StartColumn!.Value));
        var linkTarget = Assert.Single(link.Metadata, metadata => metadata.Name == "target");
        var linkTitle = Assert.Single(link.Metadata, metadata => metadata.Name == "title");
        Assert.Equal(new MarkdownSourceSpan(1, 25, 1, 43), linkTarget.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 46, 1, 49), linkTitle.SourceSpan);
        Assert.Equal(
            "# Native **AST** [docs](https://contoso.test \"Docs\")",
            native.CreateReplaceEdit(linkTarget, "https://contoso.test").Apply(native.SourceMarkdown).Split('\n')[0]);
        Assert.Equal(
            "# Native **AST** [docs](https://example.com \"Guide\")",
            native.CreateReplaceEdit(linkTitle, "Guide").Apply(native.SourceMarkdown).Split('\n')[0]);

        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(native.Blocks[1]);
        Assert.Contains(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Code && inline.Text == "code");
        var footnoteRef = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.FootnoteRef);
        Assert.Equal("note", footnoteRef.GetMetadata("label"));
        var footnoteLabel = Assert.Single(footnoteRef.Metadata, metadata => metadata.Name == "label");
        Assert.Equal(new MarkdownSourceSpan(3, 34, 3, 37), footnoteLabel.SourceSpan);
        Assert.Equal(
            "Paragraph with `code`, footnote[^memo], and ![Alt](img.png \"Img\").",
            native.CreateReplaceEdit(footnoteLabel, "memo").Apply(native.SourceMarkdown).Split('\n')[2]);
        var image = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Image);
        Assert.Equal("Alt", image.GetMetadata("alt"));
        Assert.Equal("img.png", image.GetMetadata("source"));
        Assert.Equal("Img", image.GetMetadata("imageTitle"));
        var imageAlt = Assert.Single(image.Metadata, metadata => metadata.Name == "alt");
        var imageSource = Assert.Single(image.Metadata, metadata => metadata.Name == "source");
        var imageTitle = Assert.Single(image.Metadata, metadata => metadata.Name == "imageTitle");
        Assert.Equal(new MarkdownSourceSpan(3, 47, 3, 49), imageAlt.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 52, 3, 58), imageSource.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 61, 3, 63), imageTitle.SourceSpan);
        Assert.Equal(
            "Paragraph with `code`, footnote[^note], and ![Logo](img.png \"Img\").",
            native.CreateReplaceEdit(imageAlt, "Logo").Apply(native.SourceMarkdown).Split('\n')[2]);
        Assert.Equal(
            "Paragraph with `code`, footnote[^note], and ![Alt](logo.svg \"Img\").",
            native.CreateReplaceEdit(imageSource, "logo.svg").Apply(native.SourceMarkdown).Split('\n')[2]);
        Assert.Equal(
            "Paragraph with `code`, footnote[^note], and ![Alt](img.png \"Diagram\").",
            native.CreateReplaceEdit(imageTitle, "Diagram").Apply(native.SourceMarkdown).Split('\n')[2]);

        var snapshotHeading = native.ToSnapshot().Blocks[0];
        Assert.Equal(1, snapshotHeading.FieldSourceSpans["level"]!.StartColumn);
        Assert.Equal(1, snapshotHeading.FieldSourceSpans["level"]!.EndColumn);
        Assert.Equal(3, snapshotHeading.FieldSourceSpans["text"]!.StartColumn);
        Assert.Equal(51, snapshotHeading.FieldSourceSpans["text"]!.EndColumn);
        var snapshotLink = Assert.Single(snapshotHeading.Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Link);
        Assert.Equal("https://example.com", snapshotLink.Metadata["target"]);
        Assert.Equal("Docs", snapshotLink.Metadata["title"]);
        Assert.Equal(25, snapshotLink.MetadataSourceSpans["target"]!.StartColumn);
        Assert.Equal(43, snapshotLink.MetadataSourceSpans["target"]!.EndColumn);
        Assert.Equal(46, snapshotLink.MetadataSourceSpans["title"]!.StartColumn);
        Assert.Equal(49, snapshotLink.MetadataSourceSpans["title"]!.EndColumn);

        var snapshotParagraph = native.ToSnapshot().Blocks[1];
        var snapshotFootnoteRef = Assert.Single(snapshotParagraph.Inlines, inline => inline.Kind == MarkdownNativeInlineKind.FootnoteRef);
        Assert.Equal("note", snapshotFootnoteRef.Metadata["label"]);
        Assert.Equal(34, snapshotFootnoteRef.MetadataSourceSpans["label"]!.StartColumn);
        Assert.Equal(37, snapshotFootnoteRef.MetadataSourceSpans["label"]!.EndColumn);
        var snapshotImage = Assert.Single(snapshotParagraph.Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Image);
        Assert.Equal("Alt", snapshotImage.Metadata["alt"]);
        Assert.Equal("img.png", snapshotImage.Metadata["source"]);
        Assert.Equal("Img", snapshotImage.Metadata["imageTitle"]);
        Assert.Equal(47, snapshotImage.MetadataSourceSpans["alt"]!.StartColumn);
        Assert.Equal(49, snapshotImage.MetadataSourceSpans["alt"]!.EndColumn);
        Assert.Equal(52, snapshotImage.MetadataSourceSpans["source"]!.StartColumn);
        Assert.Equal(58, snapshotImage.MetadataSourceSpans["source"]!.EndColumn);
        Assert.Equal(61, snapshotImage.MetadataSourceSpans["imageTitle"]!.StartColumn);
        Assert.Equal(63, snapshotImage.MetadataSourceSpans["imageTitle"]!.EndColumn);
    }

    [Fact]
    public void Parse_Restricts_BlockFirst_List_Item_InlineRuns_To_Lead_Content() {
        const string markdown = """
- - foo
- # Bar
""";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        var list = Assert.IsType<MarkdownNativeListBlock>(Assert.Single(native.Blocks));
        Assert.Equal(2, list.Items.Count);
        Assert.Empty(list.Items[0].Text);
        Assert.Empty(list.Items[0].InlineRuns);
        Assert.Empty(list.Items[1].Text);
        Assert.Empty(list.Items[1].InlineRuns);

        var nestedList = Assert.IsType<MarkdownNativeListBlock>(Assert.Single(list.Items[0].Children));
        var nestedItem = Assert.Single(nestedList.Items);
        var nestedInline = Assert.Single(nestedItem.InlineRuns);
        Assert.Equal("foo", nestedInline.Text);

        var heading = Assert.IsType<MarkdownNativeHeadingBlock>(Assert.Single(list.Items[1].Children));
        var headingInline = Assert.Single(heading.InlineRuns);
        Assert.Equal("Bar", headingInline.Text);

        Assert.Single(native.EnumerateInlines(), inline => inline.Text == "foo");
        Assert.Single(native.EnumerateInlines(), inline => inline.Text == "Bar");
    }

    [Fact]
    public void Parse_Includes_Syntax_Path_In_Block_Ids_For_Duplicated_Generated_Blocks() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new DuplicateParagraphTransform());

        var native = MarkdownNativeDocument.Parse("Same", options);

        Assert.Equal(2, native.Blocks.Count);
        Assert.All(native.Blocks, block => Assert.Equal(MarkdownNativeBlockKind.Paragraph, block.Kind));
        Assert.Equal(2, native.Blocks.Select(block => block.Id).Distinct().Count());
        Assert.Equal(native.Blocks[0], native.FindBlockById(native.Blocks[0].Id));
        Assert.Equal(native.Blocks[1], native.FindBlockById(native.Blocks[1].Id));

        var inlines = native.EnumerateInlines().Where(inline => inline.Text == "Same").ToArray();
        Assert.Equal(2, inlines.Length);
        Assert.Equal(2, inlines.Select(inline => inline.Id).Distinct().Count());
        Assert.Equal(inlines[0], native.FindInlineById(inlines[0].Id));
        Assert.Equal(inlines[1], native.FindInlineById(inlines[1].Id));
    }

    [Fact]
    public void SyntaxNode_IndexInParent_Does_Not_Rescan_Wide_Sibling_Lists() {
        var childArray = Enumerable.Range(0, 2048)
            .Select(_ => new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineText, literal: "x"))
            .ToArray();
        var children = new CountingReadOnlyList<MarkdownSyntaxNode>(childArray);

        _ = new MarkdownSyntaxNode(MarkdownSyntaxKind.Paragraph, children: children);
        children.ResetIndexerReads();

        for (var i = 0; i < childArray.Length; i++) {
            Assert.Equal(i, childArray[i].IndexInParent);
        }

        Assert.Equal(0, children.IndexerReads);
    }

    [Fact]
    public void Parse_Exposes_Visual_Payload_Hints_Without_Json_Dependency() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence));
        options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            "Mermaid",
            new[] { "mermaid" },
            context => new SemanticFencedBlock(MarkdownSemanticKinds.Mermaid, context.InfoString, context.Content, context.Caption)));
        var markdown = """
```ix-chart {#cpu .wide title="CPU" pinned=true}
{"type":"bar","data":{"labels":["A"],"datasets":[{"label":"CPU","data":[42]}]}}
```

```mermaid
graph TD
  A-->B
```
""";

        var native = MarkdownNativeDocument.Parse(markdown, options);

        var chart = Assert.IsType<MarkdownNativeVisualBlock>(native.Blocks[0]);
        Assert.Equal(MarkdownNativeVisualPayloadFormat.JsonObject, chart.Payload.Format);
        Assert.True(chart.Payload.IsJson);
        Assert.Equal(MarkdownSemanticKinds.Chart, chart.Payload.DetectedSemanticKind);
        Assert.Equal("bar", chart.Payload.JsonType);
        Assert.Equal("true", chart.Payload.Signals["json.has.data"]);

        var mermaid = Assert.IsType<MarkdownNativeVisualBlock>(native.Blocks[1]);
        Assert.Equal(MarkdownNativeVisualPayloadFormat.Mermaid, mermaid.Payload.Format);
        Assert.True(mermaid.Payload.IsMermaid);
    }

    [Fact]
    public void ToSnapshot_Returns_UI_Safe_Block_Inline_Table_And_Diagnostic_Data() {
        var markdown = """
# Snapshot [docs](https://example.com)

| Name | Value |
| :--- | ---: |
| **CPU** | `42` |

***
""";

        var native = MarkdownNativeDocument.Parse(markdown);

        var snapshot = native.ToSnapshot();

        Assert.Equal(MarkdownNativeDocumentSourceKind.ReaderInput, snapshot.SourceKind);
        Assert.Equal(native.Blocks.Count, snapshot.Blocks.Count);
        var heading = snapshot.Blocks[0];
        Assert.Equal(MarkdownNativeBlockKind.Heading, heading.Kind);
        Assert.Equal("Snapshot docs", heading.Text);
        Assert.Contains(heading.Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Link && inline.Metadata["target"] == "https://example.com");

        var table = snapshot.Blocks[1];
        Assert.Equal(ColumnAlignment.Left, table.HeaderCells[0].Alignment);
        Assert.Equal(ColumnAlignment.Right, table.Rows[0][1].Alignment);
        Assert.Contains(table.Rows[0][0].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Strong && inline.Text == "CPU");

        var thematicBreak = snapshot.Blocks[2];
        Assert.Equal(MarkdownNativeBlockKind.ThematicBreak, thematicBreak.Kind);
        Assert.Equal("---", thematicBreak.Markdown);
        Assert.DoesNotContain(snapshot.Diagnostics, diagnostic =>
            diagnostic.Id == "native.unsupported-block"
            && diagnostic.BlockId == native.Blocks[2].Id);
    }

    [Fact]
    public void Navigation_And_Source_Edit_Helpers_Find_Blocks_And_Draft_Replacements() {
        var markdown = """
> [!NOTE] Heads up
> Body text

Outside
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var callout = Assert.IsType<MarkdownNativeCalloutBlock>(native.Blocks[0]);
        var child = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(callout.Children));
        var snapshot = native.ToSnapshot().Blocks[0];

        Assert.Same(callout, native.FindBlockById(callout.Id));
        Assert.Same(child, native.FindBlockAtPosition(child.SourceSpan!.Value.StartLine, child.SourceSpan.Value.StartColumn!.Value));
        Assert.Equal(new[] { callout.Id, child.Id }, native.GetBlockPath(child.Id).Select(block => block.Id).ToArray());
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 4), callout.OpeningMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 8), callout.KindSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 9, 1, 9), callout.ClosingMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 1, 18), callout.TitleSourceSpan);
        Assert.Equal(3, snapshot.FieldSourceSpans["calloutOpeningMarker"]!.StartColumn);
        Assert.Equal(4, snapshot.FieldSourceSpans["calloutOpeningMarker"]!.EndColumn);
        Assert.Equal(5, snapshot.FieldSourceSpans["calloutKind"]!.StartColumn);
        Assert.Equal(8, snapshot.FieldSourceSpans["calloutKind"]!.EndColumn);
        Assert.Equal(9, snapshot.FieldSourceSpans["calloutClosingMarker"]!.StartColumn);
        Assert.Equal(9, snapshot.FieldSourceSpans["calloutClosingMarker"]!.EndColumn);
        Assert.Equal(11, snapshot.FieldSourceSpans["title"]!.StartColumn);
        Assert.Equal(18, snapshot.FieldSourceSpans["title"]!.EndColumn);

        var withTitle = native.CreateReplaceEdit(callout.TitleSourceSpan!.Value, "New title").Apply(native.SourceMarkdown);
        Assert.Contains("> [!NOTE] New title", withTitle);
        Assert.DoesNotContain("> [!NOTE] Heads up", withTitle);

        var edit = native.CreateReplaceEdit(child, "> Updated body");
        var updated = edit.Apply(native.SourceMarkdown);
        Assert.Contains("> Updated body", updated);
        Assert.DoesNotContain("> Body text", updated);
    }

    [Fact]
    public void Source_Edit_Helpers_Use_Normalized_Source_That_Backs_SourceSpans() {
        var markdown = "First\r\n\r\nSecond\r\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(native.Blocks[1]);

        Assert.Equal("First\n\nSecond\n", native.SourceMarkdown);
        var edit = native.CreateReplaceEdit(paragraph, "Updated");
        var updated = edit.Apply(native.SourceMarkdown);

        Assert.Equal("First\n\nUpdated\n", updated);
    }

    [Fact]
    public void Source_Edit_Helpers_Replace_Fenced_Code_Block_And_Preserve_Surrounding_Source() {
        var markdown = """
Before

```csharp
Console.WriteLine("old");
```

After
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var code = Assert.IsType<MarkdownNativeCodeBlock>(native.Blocks[1]);

        var edit = native.CreateReplaceEdit(code, """
```csharp
Console.WriteLine("new");
```
""");
        var updated = edit.Apply(native.SourceMarkdown);

        Assert.Equal(NormalizeLineEndings("""
Before

```csharp
Console.WriteLine("new");
```

After
"""), NormalizeLineEndings(updated));
    }

    [Fact]
    public void Source_Edit_Helpers_Replace_Fenced_Code_Info_And_Content_Tokens() {
        var markdown = """
Before

```csharp
Console.WriteLine("old");
```

After
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var code = Assert.IsType<MarkdownNativeCodeBlock>(native.Blocks[1]);

        var withInfo = native.CreateReplaceEdit(code.InfoStringSourceSpan!.Value, "powershell").Apply(native.SourceMarkdown);
        Assert.Equal(NormalizeLineEndings("""
Before

```powershell
Console.WriteLine("old");
```

After
"""), NormalizeLineEndings(withInfo));

        var withContent = native.CreateReplaceEdit(code.ContentSourceSpan!.Value, "Write-Host \"new\"").Apply(native.SourceMarkdown);
        Assert.Equal(NormalizeLineEndings("""
Before

```csharp
Write-Host "new"
```

After
"""), NormalizeLineEndings(withContent));
    }

    [Fact]
    public void Source_Edit_Helpers_Replace_Heading_Level_And_Text_Tokens() {
        var markdown = """
# Old **Title**

Body
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var heading = Assert.IsType<MarkdownNativeHeadingBlock>(native.Blocks[0]);

        var withLevel = native.CreateReplaceEdit(heading.LevelSourceSpan!.Value, "##").Apply(native.SourceMarkdown);
        Assert.Equal(NormalizeLineEndings("""
## Old **Title**

Body
"""), NormalizeLineEndings(withLevel));

        var withText = native.CreateReplaceEdit(heading.TextSourceSpan!.Value, "New Title").Apply(native.SourceMarkdown);
        Assert.Equal(NormalizeLineEndings("""
# New Title

Body
"""), NormalizeLineEndings(withText));
    }

    [Fact]
    public void Source_Edit_Helpers_Replace_Inline_Token_And_Preserve_Surrounding_Source() {
        var markdown = """
Before `old` and **bold** after.

Next paragraph.
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var codeInline = Assert.Single(
            native.EnumerateInlines(),
            inline => inline.Kind == MarkdownNativeInlineKind.Code && inline.Text == "old");

        var edit = native.CreateReplaceEdit(codeInline, "`new`");
        var updated = edit.Apply(native.SourceMarkdown);

        Assert.Equal(NormalizeLineEndings("""
Before `new` and **bold** after.

Next paragraph.
"""), NormalizeLineEndings(updated));
    }

    [Fact]
    public void Renderer_ParseNativeDocument_Uses_Reader_Normalized_Source_For_Edits() {
        var markdown = "First\r\n\r\nSecond\r\n";

        var parseResult = OfficeIMO.MarkdownRenderer.MarkdownRenderer.ParseDocumentResult(markdown);
        var native = OfficeIMO.MarkdownRenderer.MarkdownRenderer.ParseNativeDocument(markdown);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(native.Blocks[1]);

        Assert.Equal("First\r\n\r\nSecond\r\n", parseResult.PreprocessedMarkdown);
        Assert.Equal("First\n\nSecond\n", parseResult.SourceMarkdown);
        Assert.Equal(MarkdownNativeDocumentSourceKind.RendererPreprocessed, native.SourceKind);
        Assert.Equal("First\n\nSecond\n", native.SourceMarkdown);

        var edit = native.CreateReplaceEdit(paragraph, "Updated");
        Assert.Equal("First\n\nUpdated\n", edit.Apply(native.SourceMarkdown));
    }

    [Fact]
    public void Source_Edit_Helpers_Reject_Spans_When_Backing_Source_Is_Missing() {
        var parseResult = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("Editable");
        var withoutSource = new MarkdownParseResult(
            parseResult.Document,
            parseResult.SyntaxTree,
            parseResult.FinalSyntaxTree,
            sourceMarkdown: null,
            transformDiagnostics: parseResult.TransformDiagnostics);
        var native = MarkdownNativeDocument.FromParseResult(withoutSource, sourceMarkdown: string.Empty);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));

        var ex = Assert.Throws<InvalidOperationException>(() => native.CreateReplaceEdit(paragraph, "Updated"));
        Assert.Contains("cannot be mapped", ex.Message);
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
    public void Parse_Projects_Transform_Diagnostic_Related_SourceSpans_Into_Native_Snapshots() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new UppercaseParagraphsTransform());

        var native = MarkdownNativeDocument.Parse("""
alpha

beta
""", options);

        var diagnostic = Assert.Single(native.Diagnostics, item => item.Id == "native.transform");
        Assert.Equal(MarkdownNativeDiagnosticSeverity.Info, diagnostic.Severity);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 4), diagnostic.SourceSpan);
        Assert.Equal(new[] { 1, 3 }, diagnostic.RelatedSourceSpans.Select(span => span.StartLine).ToArray());
        Assert.Equal(new[] { 1, 3 }, diagnostic.RelatedSourceSpans.Select(span => span.EndLine).ToArray());

        var snapshotDiagnostic = Assert.Single(native.ToSnapshot().Diagnostics, item => item.Id == "native.transform");
        Assert.Equal(1, snapshotDiagnostic.SourceSpan!.StartLine);
        Assert.Equal(3, snapshotDiagnostic.SourceSpan.EndLine);
        Assert.Equal(new[] { 1, 3 }, snapshotDiagnostic.RelatedSourceSpans.Select(span => span.StartLine).ToArray());
        Assert.Equal(new[] { 1, 3 }, snapshotDiagnostic.RelatedSourceSpans.Select(span => span.EndLine).ToArray());
    }

    [Fact]
    public void Parse_Projects_Transform_Diagnostic_Precise_Node_SourceSpan_Into_Native_Snapshots() {
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.DocumentTransforms.Add(new MarkdownInlineNormalizationTransform(new MarkdownInputNormalizationOptions {
            NormalizeTightStrongBoundaries = true
        }));

        var native = MarkdownNativeDocument.Parse("Prefix **bold**suffix", options);

        var diagnostic = Assert.Single(native.Diagnostics, item => item.Id == "native.transform");
        Assert.Equal(new MarkdownSourceSpan(1, 16, 1, 21), diagnostic.SourceSpan);

        var snapshotDiagnostic = Assert.Single(native.ToSnapshot().Diagnostics, item => item.Id == "native.transform");
        Assert.Equal(1, snapshotDiagnostic.SourceSpan!.StartLine);
        Assert.Equal(16, snapshotDiagnostic.SourceSpan.StartColumn);
        Assert.Equal(1, snapshotDiagnostic.SourceSpan.EndLine);
        Assert.Equal(21, snapshotDiagnostic.SourceSpan.EndColumn);
    }

    [Fact]
    public void Parse_Projects_Nested_Transform_Diagnostic_Precise_Node_SourceSpan_Into_Native_Snapshots() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownInlineNormalizationTransform(new MarkdownInputNormalizationOptions {
            NormalizeTightColonSpacing = true
        }));

        var native = MarkdownNativeDocument.Parse("""
> [!NOTE] Why it matters
> coverage:missing evidence
""", options);

        var diagnostic = Assert.Single(native.Diagnostics, item => item.Id == "native.transform");
        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 27), diagnostic.SourceSpan);

        var snapshotDiagnostic = Assert.Single(native.ToSnapshot().Diagnostics, item => item.Id == "native.transform");
        Assert.Equal(2, snapshotDiagnostic.SourceSpan!.StartLine);
        Assert.Equal(3, snapshotDiagnostic.SourceSpan.StartColumn);
        Assert.Equal(2, snapshotDiagnostic.SourceSpan.EndLine);
        Assert.Equal(27, snapshotDiagnostic.SourceSpan.EndColumn);
    }

    [Fact]
    public void Parse_Reports_Generated_Definition_Child_Diagnostics_For_Rebuilt_Paragraph_Wrappers() {
        var options = new MarkdownReaderOptions { PreserveTrivia = true };
        options.DocumentTransforms.Add(new RewriteFirstDefinitionBodyTransform("generated"));

        var native = MarkdownNativeDocument.Parse("Term: original", options);

        var definitionList = Assert.IsType<MarkdownNativeDefinitionListBlock>(Assert.Single(native.Blocks));
        var definition = Assert.Single(Assert.Single(definitionList.Groups).Definitions);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(definition.Children));
        var diagnostic = Assert.Single(native.Diagnostics, item => item.Id == "native.generated-definition-child");

        Assert.Equal("generated", paragraph.Text);
        Assert.True(paragraph.SyntaxNode.IsGenerated);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 14), paragraph.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 14), diagnostic.SourceSpan);
        Assert.Same(paragraph, diagnostic.Block);
        Assert.Equal(MarkdownNativeDiagnosticSeverity.Info, diagnostic.Severity);

        var generatedDiagnostics = native.ParseResult.GeneratedSyntaxDiagnostics;
        var generatedParagraph = Assert.Single(generatedDiagnostics, item => item.Kind == MarkdownSyntaxKind.Paragraph);
        Assert.Equal("syntax.generated-node", generatedParagraph.Id);
        Assert.Equal("Document > DefinitionList > DefinitionGroup > DefinitionValue > Paragraph", generatedParagraph.SyntaxPath);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 14), generatedParagraph.SourceSpan);
        Assert.Equal(nameof(ParagraphBlock), generatedParagraph.AssociatedObjectType);
        Assert.Same(paragraph.SyntaxNode, generatedParagraph.SyntaxNode);
        Assert.Same(paragraph.SourceBlock, generatedParagraph.AssociatedObject);

        Assert.False(native.ParseResult.TryCreateOriginalSourceSlice(paragraph.SyntaxNode, out _, out var parseFailureReason));
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.GeneratedSyntaxNode, parseFailureReason);
        Assert.False(native.TryCreateOriginalSourceSlice(paragraph, out _, out var nativeFailureReason));
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.GeneratedSyntaxNode, nativeFailureReason);

        var edit = native.CreateReplaceEdit(paragraph, "updated");
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.GeneratedSyntaxNode, edit.OriginalSourceFailureReason);
        Assert.Equal("Term: updated", edit.Apply(native.SourceMarkdown));

        var roundtrip = native.WriteWithSourceEdit(edit);
        Assert.Equal("Term: updated", roundtrip.Markdown);
        Assert.Contains(roundtrip.Diagnostics, item => item.Id == "roundtrip.document-transformed");
        var sourceFailure = Assert.Single(
            roundtrip.Diagnostics,
            item => item.Id == "roundtrip.original-source-slice-unavailable");
        Assert.Contains("generated from semantic content", sourceFailure.Message, StringComparison.Ordinal);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 14), sourceFailure.SourceSpan);
    }

    [Fact]
    public void Parse_Reports_Fallback_Diagnostics_For_Unsupported_Blocks() {
        var native = MarkdownNativeDocument.Parse("[TOC]");

        var other = Assert.IsType<MarkdownNativeOtherBlock>(Assert.Single(native.Blocks));
        var diagnostic = Assert.Single(native.Diagnostics, item => item.Id == "native.unsupported-block");
        Assert.Same(other, diagnostic.Block);
        Assert.Equal(MarkdownNativeDiagnosticSeverity.Info, diagnostic.Severity);
    }

    private sealed class DuplicateParagraphTransform : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            var clone = MarkdownDoc.Create();
            clone.Add(new ParagraphBlock(new InlineSequence().Text("Same")));
            clone.Add(new ParagraphBlock(new InlineSequence().Text("Same")));
            return clone;
        }
    }

    private sealed class UppercaseParagraphsTransform : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            var transformed = MarkdownDoc.Create();
            foreach (var block in document.Blocks) {
                if (block is ParagraphBlock paragraph) {
                    transformed.Add(new ParagraphBlock(new InlineSequence().Text(paragraph.Inlines.RenderMarkdown().ToUpperInvariant())));
                } else {
                    transformed.Add(block);
                }
            }

            return transformed;
        }
    }

    private sealed class RewriteFirstDefinitionBodyTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            var definitionList = document.Blocks.OfType<DefinitionListBlock>().FirstOrDefault();
            var definition = definitionList?.Groups.FirstOrDefault()?.Definitions.FirstOrDefault();
            if (definition == null) {
                return document;
            }

            definition.Blocks.Clear();
            definition.Blocks.Add(new ParagraphBlock(new InlineSequence().Text(text)));
            return document;
        }
    }

    private static string NormalizeLineEndings(string value) => value.Replace("\r\n", "\n");

    private sealed class CountingReadOnlyList<T> : IReadOnlyList<T> {
        private readonly IReadOnlyList<T> _items;

        internal CountingReadOnlyList(IReadOnlyList<T> items) {
            _items = items;
        }

        internal int IndexerReads { get; private set; }

        public int Count => _items.Count;

        public T this[int index] {
            get {
                IndexerReads++;
                return _items[index];
            }
        }

        public IEnumerator<T> GetEnumerator() => _items.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        internal void ResetIndexerReads() {
            IndexerReads = 0;
        }
    }
}
