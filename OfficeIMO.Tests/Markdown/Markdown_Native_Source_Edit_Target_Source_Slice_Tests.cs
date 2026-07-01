using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Native_Source_Edit_Target_Source_Slice_Tests {
    [Fact]
    public void NativeBlock_SourceField_Snapshots_Expose_Normalized_And_Original_Source_Text() {
        var native = MarkdownNativeDocument.Parse(
            "# Title\r\n\r\n> [!NOTE] Heads up\r\n> Body",
            new MarkdownReaderOptions { PreserveTrivia = true });

        var snapshot = native.ToSnapshot();
        var headingText = Assert.Single(snapshot.Blocks[0].SourceFields, field => field.Name == "text");
        var calloutTitle = Assert.Single(snapshot.Blocks[1].SourceFields, field => field.Name == "calloutTitle");
        var calloutBody = Assert.Single(snapshot.Blocks[1].SourceFields, field => field.Name == "calloutBody");

        Assert.Equal("Title", headingText.SourceText);
        Assert.Equal("Title", headingText.OriginalSourceText);
        Assert.Null(headingText.OriginalSourceFailureReason);

        Assert.Equal("Heads up", calloutTitle.SourceText);
        Assert.Equal("Heads up", calloutTitle.OriginalSourceText);
        Assert.Null(calloutTitle.OriginalSourceFailureReason);

        Assert.Equal("Body", calloutBody.SourceText);
        Assert.Equal("Body", calloutBody.OriginalSourceText);
        Assert.Null(calloutBody.OriginalSourceFailureReason);
    }

    [Fact]
    public void NativeBlock_SourceField_Snapshots_Report_Original_Source_Failure_When_Trivia_Is_Not_Preserved() {
        var native = MarkdownNativeDocument.Parse("# Title");

        var snapshot = native.ToSnapshot();
        var headingText = Assert.Single(snapshot.Blocks[0].SourceFields, field => field.Name == "text");

        Assert.Equal("Title", headingText.SourceText);
        Assert.Null(headingText.OriginalSourceText);
        Assert.Equal(
            MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved,
            headingText.OriginalSourceFailureReason);
    }

    [Fact]
    public void NativeInline_Metadata_Snapshots_Expose_Normalized_And_Original_Source_Text() {
        var native = MarkdownNativeDocument.Parse(
            "See [docs](https://example.com \"Docs\") and \\* &copy;",
            new MarkdownReaderOptions { PreserveTrivia = true });

        var snapshot = native.ToSnapshot();
        var link = Assert.Single(snapshot.Blocks[0].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var openingMarker = Assert.Single(link.MetadataFields, field => field.Name == "openingMarker");
        var target = Assert.Single(link.MetadataFields, field => field.Name == "target");
        var title = Assert.Single(link.MetadataFields, field => field.Name == "title");
        var closingMarker = Assert.Single(link.MetadataFields, field => field.Name == "closingMarker");
        var inlineMetadata = snapshot.Blocks[0].Inlines.SelectMany(inline => inline.MetadataFields).ToArray();
        var escapeMarker = Assert.Single(inlineMetadata, field => field.Name == "escapeMarker");
        var entitySourceText = Assert.Single(inlineMetadata, field => field.Name == "sourceText");

        Assert.Equal("[", openingMarker.SourceText);
        Assert.Equal("[", openingMarker.OriginalSourceText);
        Assert.Null(openingMarker.OriginalSourceFailureReason);

        Assert.Equal("https://example.com", target.SourceText);
        Assert.Equal("https://example.com", target.OriginalSourceText);
        Assert.Null(target.OriginalSourceFailureReason);

        Assert.Equal("Docs", title.SourceText);
        Assert.Equal("Docs", title.OriginalSourceText);
        Assert.Null(title.OriginalSourceFailureReason);

        Assert.Equal(")", closingMarker.SourceText);
        Assert.Equal(")", closingMarker.OriginalSourceText);
        Assert.Null(closingMarker.OriginalSourceFailureReason);

        Assert.Equal("\\", escapeMarker.SourceText);
        Assert.Equal("\\", escapeMarker.OriginalSourceText);
        Assert.Null(escapeMarker.OriginalSourceFailureReason);

        Assert.Equal("&copy;", entitySourceText.SourceText);
        Assert.Equal("&copy;", entitySourceText.OriginalSourceText);
        Assert.Null(entitySourceText.OriginalSourceFailureReason);
    }

    [Fact]
    public void NativeInline_Metadata_Snapshots_Report_Original_Source_Failure_When_Trivia_Is_Not_Preserved() {
        var native = MarkdownNativeDocument.Parse("[docs](https://example.com)");

        var snapshot = native.ToSnapshot();
        var link = Assert.Single(snapshot.Blocks[0].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var target = Assert.Single(link.MetadataFields, field => field.Name == "target");

        Assert.Equal("https://example.com", target.SourceText);
        Assert.Null(target.OriginalSourceText);
        Assert.Equal(
            MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved,
            target.OriginalSourceFailureReason);
    }

    [Fact]
    public void NativeBlock_And_ListItem_SourceSlices_Match_ReplaceEdit_Targets() {
        var native = MarkdownNativeDocument.Parse("""
# Title

- Plain
""");
        var heading = Assert.IsType<MarkdownNativeHeadingBlock>(native.Blocks[0]);
        var item = Assert.Single(native.EnumerateListItems());

        Assert.True(native.TryCreateSourceSlice(heading, out var headingSlice));
        Assert.True(native.TryCreateSourceSlice(item, out var itemSlice));

        Assert.Equal("# Title", headingSlice.Text);
        Assert.Equal("Plain", itemSlice.Text);
        Assert.Contains("- Updated", native.CreateReplaceEdit(item, "Updated").Apply(native.SourceMarkdown));
    }

    [Fact]
    public void NativeListItem_Paragraph_SourceSlices_Address_Individual_Paragraphs() {
        var native = MarkdownNativeDocument.Parse("""
- alpha **one**

  beta [two](https://example.com)
""", new MarkdownReaderOptions { PreserveTrivia = true });
        var item = Assert.Single(native.EnumerateListItems());

        var paragraphs = item.Paragraphs.ToArray();

        Assert.Equal(2, paragraphs.Length);
        Assert.Equal("alpha one", paragraphs[0].Text);
        Assert.Equal("beta two", paragraphs[1].Text);
        Assert.True(native.TryCreateSourceSlice(paragraphs[0], out var firstSlice));
        Assert.True(native.TryCreateSourceSlice(paragraphs[1], out var secondSlice));
        Assert.Equal("alpha **one**", firstSlice.Text);
        Assert.Equal("beta [two](https://example.com)", secondSlice.Text);
        Assert.NotNull(paragraphs[1].SyntaxNode);
        Assert.False(paragraphs[1].SyntaxNode.IsGenerated);
        Assert.True(native.TryCreateOriginalSourceSlice(paragraphs[1], out var originalSecondSlice, out var originalReason));
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.None, originalReason);
        Assert.Equal("beta [two](https://example.com)", originalSecondSlice.Text);
        Assert.Contains(paragraphs[0].InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Strong && inline.Text == "one");
        Assert.Contains(paragraphs[1].InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Link && inline.Text == "two");

        var snapshotParagraphs = Assert.Single(native.ToSnapshot().Blocks).Items[0].Paragraphs;
        Assert.Equal(2, snapshotParagraphs.Count);
        Assert.Equal("alpha one", snapshotParagraphs[0].Text);
        Assert.Equal("beta two", snapshotParagraphs[1].Text);
        Assert.NotNull(snapshotParagraphs[0].SourceSpan);
        Assert.NotNull(snapshotParagraphs[1].SourceSpan);
        Assert.Contains(snapshotParagraphs[0].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Strong && inline.Text == "one");
        Assert.Contains(snapshotParagraphs[1].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Link && inline.Text == "two");

        var edit = native.CreateReplaceEdit(paragraphs[1], "gamma [three](https://example.org)");
        var roundtrip = native.WriteWithSourceEdit(edit);

        Assert.Contains("  gamma [three](https://example.org)", edit.Apply(native.SourceMarkdown));
        Assert.Contains("  gamma [three](https://example.org)", roundtrip.Markdown);
        Assert.Empty(roundtrip.Diagnostics);
    }

    [Fact]
    public void NativeDocument_SourceTrivia_SourceSlices_Address_Blank_And_Horizontal_Whitespace() {
        var native = MarkdownNativeDocument.Parse(
            "# Title  \r\n\tIndented text\t \r\n   \r\nParagraph",
            new MarkdownReaderOptions { PreserveTrivia = true });

        var trivia = native.SourceTrivia.Where(item => item.Kind != MarkdownNativeSourceTriviaKind.LineEnding).ToArray();
        var trailingTrivia = native.EnumerateSourceTrivia(MarkdownNativeSourceTriviaKind.TrailingWhitespace).ToArray();
        var leadingTrivia = Assert.Single(native.EnumerateSourceTrivia(MarkdownNativeSourceTriviaKind.LeadingWhitespace));
        var blankLineTrivia = Assert.Single(native.EnumerateSourceTrivia(MarkdownNativeSourceTriviaKind.BlankLine));

        Assert.Equal(4, trivia.Length);
        Assert.Equal(new[] {
            MarkdownNativeSourceTriviaKind.TrailingWhitespace,
            MarkdownNativeSourceTriviaKind.LeadingWhitespace,
            MarkdownNativeSourceTriviaKind.TrailingWhitespace,
            MarkdownNativeSourceTriviaKind.BlankLine
        }, trivia.Select(item => item.Kind).ToArray());
        Assert.Equal(2, trailingTrivia.Length);
        Assert.Equal("  ", trailingTrivia[0].Text);
        Assert.Equal("\t", leadingTrivia.Text);
        Assert.Equal("\t ", trailingTrivia[1].Text);
        Assert.Equal("   ", blankLineTrivia.Text);
        Assert.Same(leadingTrivia, native.FindSourceTriviaAtPosition(2, 1));
        Assert.Same(trailingTrivia[1], native.FindSourceTriviaAtPosition(2, trailingTrivia[1].SourceSpan.StartColumn!.Value));
        Assert.Same(blankLineTrivia, native.FindSourceTriviaAtPosition(3, 1));
        Assert.True(native.TryCreateSourceSlice(trailingTrivia[0], out var titleTrailingSlice));
        Assert.True(native.TryCreateSourceSlice(leadingTrivia, out var leadingSlice));
        Assert.True(native.TryCreateSourceSlice(trailingTrivia[1], out var indentedTrailingSlice));
        Assert.True(native.TryCreateSourceSlice(blankLineTrivia, out var blankLineSlice));
        Assert.Equal("  ", titleTrailingSlice.Text);
        Assert.Equal("\t", leadingSlice.Text);
        Assert.Equal("\t ", indentedTrailingSlice.Text);
        Assert.Equal("   ", blankLineSlice.Text);
        Assert.True(native.TryCreateOriginalSourceSlice(leadingTrivia, out var originalLeadingSlice, out var leadingReason));
        Assert.True(native.TryCreateOriginalSourceSlice(trailingTrivia[1], out var originalTrailingSlice, out var trailingReason));
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.None, leadingReason);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.None, trailingReason);
        Assert.Equal("\t", originalLeadingSlice.Text);
        Assert.Equal("\t ", originalTrailingSlice.Text);

        var snapshotTrivia = native.ToSnapshot().SourceTrivia
            .Where(item => item.Kind != MarkdownNativeSourceTriviaKind.LineEnding)
            .ToArray();
        Assert.Equal(4, snapshotTrivia.Length);
        Assert.Equal(MarkdownNativeSourceTriviaKind.TrailingWhitespace, snapshotTrivia[0].Kind);
        Assert.Equal("  ", snapshotTrivia[0].Text);
        Assert.Equal("  ", snapshotTrivia[0].SourceText);
        Assert.Equal("  ", snapshotTrivia[0].OriginalSourceText);
        Assert.Null(snapshotTrivia[0].OriginalSourceFailureReason);
        Assert.Equal(MarkdownNativeSourceTriviaKind.LeadingWhitespace, snapshotTrivia[1].Kind);
        Assert.Equal("\t", snapshotTrivia[1].Text);
        Assert.Equal("\t", snapshotTrivia[1].SourceText);
        Assert.Equal("\t", snapshotTrivia[1].OriginalSourceText);
        Assert.Null(snapshotTrivia[1].OriginalSourceFailureReason);
        Assert.Equal(MarkdownNativeSourceTriviaKind.TrailingWhitespace, snapshotTrivia[2].Kind);
        Assert.Equal("\t ", snapshotTrivia[2].SourceText);
        Assert.Equal("\t ", snapshotTrivia[2].OriginalSourceText);
        Assert.Null(snapshotTrivia[2].OriginalSourceFailureReason);
        Assert.Equal(MarkdownNativeSourceTriviaKind.BlankLine, snapshotTrivia[3].Kind);
        Assert.Equal("   ", snapshotTrivia[3].Text);
        Assert.Equal("   ", snapshotTrivia[3].SourceText);
        Assert.Equal("   ", snapshotTrivia[3].OriginalSourceText);
        Assert.Null(snapshotTrivia[3].OriginalSourceFailureReason);
    }

    [Fact]
    public void NativeDocument_SourceTrivia_SourceSlices_Address_Line_Endings_With_Original_Mapping() {
        var native = MarkdownNativeDocument.Parse(
            "# Title\r\nParagraph\rTail\nEnd",
            new MarkdownReaderOptions { PreserveTrivia = true });

        var lineEndings = native.EnumerateSourceTrivia(MarkdownNativeSourceTriviaKind.LineEnding).ToArray();

        Assert.Equal(3, lineEndings.Length);
        Assert.Equal(new[] { 8, 10, 5 }, lineEndings.Select(item => item.SourceSpan.StartColumn!.Value).ToArray());
        Assert.All(lineEndings, item => Assert.Equal("\n", item.Text));
        Assert.Same(lineEndings[0], native.FindSourceTriviaAtPosition(1, 8));
        Assert.Same(lineEndings[1], native.FindSourceTriviaAtPosition(2, 10));
        Assert.Same(lineEndings[2], native.FindSourceTriviaAtPosition(3, 5));

        var expectedOriginalText = new[] { "\r\n", "\r", "\n" };
        for (var i = 0; i < lineEndings.Length; i++) {
            Assert.True(native.TryCreateSourceSlice(lineEndings[i], out var normalizedSlice));
            Assert.True(native.TryCreateOriginalSourceSlice(lineEndings[i], out var originalSlice, out var failureReason));
            Assert.Equal(MarkdownOriginalSourceSliceFailureReason.None, failureReason);
            Assert.Equal("\n", normalizedSlice.Text);
            Assert.Equal(expectedOriginalText[i], originalSlice.Text);
        }

        var snapshotTrivia = native.ToSnapshot().SourceTrivia
            .Where(item => item.Kind == MarkdownNativeSourceTriviaKind.LineEnding)
            .ToArray();
        Assert.Equal(3, snapshotTrivia.Length);
        Assert.All(snapshotTrivia, item => Assert.Equal("\n", item.Text));
        Assert.All(snapshotTrivia, item => Assert.Equal("\n", item.SourceText));
        Assert.Equal(expectedOriginalText, snapshotTrivia.Select(item => item.OriginalSourceText).ToArray());
        Assert.All(snapshotTrivia, item => Assert.Null(item.OriginalSourceFailureReason));
    }

    [Fact]
    public void NativeDocument_SourceTrivia_Snapshots_Report_Original_Source_Failure_When_Trivia_Is_Not_Preserved() {
        var native = MarkdownNativeDocument.Parse("# Title  \n");

        var snapshotTrivia = native.ToSnapshot().SourceTrivia.ToArray();

        Assert.Equal(2, snapshotTrivia.Length);
        Assert.Equal(MarkdownNativeSourceTriviaKind.TrailingWhitespace, snapshotTrivia[0].Kind);
        Assert.Equal("  ", snapshotTrivia[0].Text);
        Assert.Equal("  ", snapshotTrivia[0].SourceText);
        Assert.Null(snapshotTrivia[0].OriginalSourceText);
        Assert.Equal(
            MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved,
            snapshotTrivia[0].OriginalSourceFailureReason);
        Assert.Equal(MarkdownNativeSourceTriviaKind.LineEnding, snapshotTrivia[1].Kind);
        Assert.Equal("\n", snapshotTrivia[1].SourceText);
        Assert.Null(snapshotTrivia[1].OriginalSourceText);
        Assert.Equal(
            MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved,
            snapshotTrivia[1].OriginalSourceFailureReason);
    }

    [Fact]
    public void NativeDocument_SourceTrivia_Line_Endings_Are_Source_Edit_Targets() {
        var native = MarkdownNativeDocument.Parse(
            "Alpha\r\nBeta\rGamma\nDelta",
            new MarkdownReaderOptions { PreserveTrivia = true });
        var lineEndings = native.EnumerateSourceTrivia(MarkdownNativeSourceTriviaKind.LineEnding).ToArray();

        var normalizedEdit = native.CreateReplaceEdit(lineEndings[1], "\n\n");
        var originalRoundtrip = native.WriteWithSourceEdit(normalizedEdit);

        Assert.Equal("Alpha\nBeta\n\nGamma\nDelta", normalizedEdit.Apply(native.SourceMarkdown));
        Assert.Equal("Alpha\r\nBeta\n\nGamma\nDelta", originalRoundtrip.Markdown);
        Assert.Empty(originalRoundtrip.Diagnostics);
    }

    [Fact]
    public void NativeDocument_SourceTrivia_Columns_Expand_Tabs_Like_Source_Map() {
        var native = MarkdownNativeDocument.Parse(
            "\tTabbed\t \n\t\nDone",
            new MarkdownReaderOptions { PreserveTrivia = true });

        var leadingTrivia = Assert.Single(native.EnumerateSourceTrivia(MarkdownNativeSourceTriviaKind.LeadingWhitespace));
        var trailingTrivia = Assert.Single(native.EnumerateSourceTrivia(MarkdownNativeSourceTriviaKind.TrailingWhitespace));
        var blankLineTrivia = Assert.Single(native.EnumerateSourceTrivia(MarkdownNativeSourceTriviaKind.BlankLine));

        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 4, 0, 0), leadingTrivia.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 1, 13, 7, 8), trailingTrivia.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 4, 10, 10), blankLineTrivia.SourceSpan);
        Assert.Same(leadingTrivia, native.FindSourceTriviaAtPosition(1, 4));
        Assert.Same(trailingTrivia, native.FindSourceTriviaAtPosition(1, 12));
        Assert.Same(trailingTrivia, native.FindSourceTriviaAtPosition(1, 13));
        Assert.Same(blankLineTrivia, native.FindSourceTriviaAtPosition(2, 4));
        Assert.True(native.TryCreateSourceSlice(leadingTrivia, out var leadingSlice));
        Assert.True(native.TryCreateSourceSlice(trailingTrivia, out var trailingSlice));
        Assert.True(native.TryCreateSourceSlice(blankLineTrivia, out var blankLineSlice));
        Assert.Equal("\t", leadingSlice.Text);
        Assert.Equal("\t ", trailingSlice.Text);
        Assert.Equal("\t", blankLineSlice.Text);
    }

    [Fact]
    public void NativeDocument_SourceEdit_Offsetless_LineColumn_Spans_Expand_Tabs_Like_Source_Map() {
        var native = MarkdownNativeDocument.Parse("\tTabbed\n");

        var edit = native.CreateReplaceEdit(new MarkdownSourceSpan(1, 5, 1, 10), "Changed");

        Assert.Equal("\tChanged\n", edit.Apply(native.SourceMarkdown));
    }

    [Fact]
    public void NativeTableCell_SourceSlice_Matches_ReplaceEdit_Target() {
        var native = MarkdownNativeDocument.Parse("""
| Name | Value |
| ---- | ----- |
| One  | Two   |
""", MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var cell = native.EnumerateTableCells().Single(candidate => !candidate.IsHeader && candidate.ColumnIndex == 1);

        Assert.True(native.TryCreateSourceSlice(cell, out var slice));

        Assert.Equal("Two", slice.Text.Trim());
        Assert.Contains("| One  | Three", native.CreateReplaceEdit(cell, "Three").Apply(native.SourceMarkdown));
    }

    [Fact]
    public void NativeDefinitionList_SourceSlices_Address_Group_Term_And_Definition_Body() {
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        options.DefinitionLists = true;
        var native = MarkdownNativeDocument.Parse("""
Term
:   First
""", options);
        var group = Assert.Single(native.EnumerateDefinitionListGroups());
        var term = Assert.Single(native.EnumerateDefinitionListTerms());
        var definition = Assert.Single(native.EnumerateDefinitionListDefinitions());

        Assert.True(native.TryCreateSourceSlice(group, out var groupSlice));
        Assert.True(native.TryCreateSourceSlice(term, out var termSlice));
        Assert.True(native.TryCreateSourceSlice(definition, out var definitionSlice));

        Assert.Equal("Term\n:   First", groupSlice.Text);
        Assert.Equal("Term", termSlice.Text);
        Assert.Equal("First", definitionSlice.Text);
    }

    [Fact]
    public void ReferenceDefinition_SourceSlices_Address_Definition_And_Fields() {
        var native = MarkdownNativeDocument.Parse("[hero]: https://example.com/docs \"Docs title\"");
        var definition = Assert.Single(native.ReferenceLinkDefinitions);
        var url = Assert.Single(native.EnumerateReferenceLinkDefinitionFields("url"));
        var title = Assert.Single(native.EnumerateReferenceLinkDefinitionFields("title"));

        Assert.True(native.TryCreateSourceSlice(definition, out var definitionSlice));
        Assert.True(native.TryCreateSourceSlice(url, out var urlSlice));
        Assert.True(native.TryCreateSourceSlice(title, out var titleSlice));

        Assert.Equal("[hero]: https://example.com/docs \"Docs title\"", definitionSlice.Text);
        Assert.Equal("https://example.com/docs", urlSlice.Text);
        Assert.Equal("Docs title", titleSlice.Text);
    }

    [Fact]
    public void ReferenceDefinition_Field_Snapshots_Expose_Normalized_And_Original_Source_Text() {
        var native = MarkdownNativeDocument.Parse(
            "[hero]: https://example.com/docs \"Docs title\"",
            new MarkdownReaderOptions { PreserveTrivia = true });

        var snapshot = Assert.Single(native.ToSnapshot().ReferenceLinkDefinitions);
        var openingMarker = Assert.Single(snapshot.SourceFields, field => field.Name == "openingMarker");
        var url = Assert.Single(snapshot.SourceFields, field => field.Name == "url");
        var title = Assert.Single(snapshot.SourceFields, field => field.Name == "title");

        Assert.Equal("[", openingMarker.SourceText);
        Assert.Equal("[", openingMarker.OriginalSourceText);
        Assert.Null(openingMarker.OriginalSourceFailureReason);

        Assert.Equal("https://example.com/docs", url.SourceText);
        Assert.Equal("https://example.com/docs", url.OriginalSourceText);
        Assert.Null(url.OriginalSourceFailureReason);

        Assert.Equal("Docs title", title.SourceText);
        Assert.Equal("Docs title", title.OriginalSourceText);
        Assert.Null(title.OriginalSourceFailureReason);
    }

    [Fact]
    public void Definition_Field_Snapshots_Report_Original_Source_Failure_When_Trivia_Is_Not_Preserved() {
        var referenceNative = MarkdownNativeDocument.Parse("[hero]: https://example.com/docs");
        var referenceSnapshot = Assert.Single(referenceNative.ToSnapshot().ReferenceLinkDefinitions);
        var referenceUrl = Assert.Single(referenceSnapshot.SourceFields, field => field.Name == "url");

        Assert.Equal("https://example.com/docs", referenceUrl.SourceText);
        Assert.Null(referenceUrl.OriginalSourceText);
        Assert.Equal(
            MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved,
            referenceUrl.OriginalSourceFailureReason);

        var abbreviationOptions = MarkdownReaderOptions.CreatePortableProfile();
        abbreviationOptions.Abbreviations = true;
        var abbreviationNative = MarkdownNativeDocument.Parse("*[HTML]: Hyper Text", abbreviationOptions);
        var abbreviationSnapshot = Assert.Single(abbreviationNative.ToSnapshot().AbbreviationDefinitions);
        var abbreviationTitle = Assert.Single(abbreviationSnapshot.SourceFields, field => field.Name == "title");

        Assert.Equal("Hyper Text", abbreviationTitle.SourceText);
        Assert.Null(abbreviationTitle.OriginalSourceText);
        Assert.Equal(
            MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved,
            abbreviationTitle.OriginalSourceFailureReason);
    }

    [Fact]
    public void OriginalSourceSlice_For_Native_SourceEdit_Targets_Returns_Failure_Reason_When_Trivia_Is_Disabled() {
        var native = MarkdownNativeDocument.Parse("- Plain");
        var item = Assert.Single(native.EnumerateListItems());

        var created = native.TryCreateOriginalSourceSlice(item, out _, out var failureReason);

        Assert.False(created);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved, failureReason);
    }
}
