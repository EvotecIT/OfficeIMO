using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Native_Source_Edit_Target_Source_Slice_Tests {
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
    public void OriginalSourceSlice_For_Native_SourceEdit_Targets_Returns_Failure_Reason_When_Trivia_Is_Disabled() {
        var native = MarkdownNativeDocument.Parse("- Plain");
        var item = Assert.Single(native.EnumerateListItems());

        var created = native.TryCreateOriginalSourceSlice(item, out _, out var failureReason);

        Assert.False(created);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved, failureReason);
    }
}
