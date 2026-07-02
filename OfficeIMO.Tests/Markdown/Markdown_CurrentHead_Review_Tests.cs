using System.Linq;
using OfficeIMO.Markdown;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_CurrentHead_Review_Tests {
    [Fact]
    public void ParseWithSyntaxTree_Preserves_TransformDiagnostics() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.PreserveTrivia = true;
        options.DocumentTransforms.Add(new MarkdownCompactHeadingBoundaryTransform());

        var result = MarkdownReader.ParseWithSyntaxTree("previous shutdown was unexpected### Reason", options);

        var diagnostic = Assert.Single(result.TransformDiagnostics);
        Assert.Contains(nameof(MarkdownCompactHeadingBoundaryTransform), diagnostic.TransformName);
    }

    [Fact]
    public void MarkdownToRtf_Does_Not_Duplicate_Loose_List_AdditionalParagraphs() {
        const string markdown = """
            - Lead

              Continuation
            """;

        var document = markdown.ToRtfDocumentFromMarkdown();

        Assert.Equal(1, document.Paragraphs.Count(paragraph => paragraph.ToPlainText() == "Continuation"));
    }

    [Fact]
    public void MarkdownToRtf_Applies_Superscript_And_Subscript_Sequence_Styles() {
        var options = new MarkdownToRtfOptions {
            ReaderOptions = MarkdownReaderOptions.CreateOfficeIMOProfile()
        };

        var document = "Formula ^2^ and H~2~O".ToRtfDocumentFromMarkdown(options);
        var paragraph = Assert.Single(document.Paragraphs);

        Assert.Contains(paragraph.Runs, run => run.Text == "2" && run.VerticalPosition == RtfVerticalPosition.Superscript);
        Assert.Contains(paragraph.Runs, run => run.Text == "2" && run.VerticalPosition == RtfVerticalPosition.Subscript);
    }

    [Fact]
    public void Sequence_RenderMarkdown_Escapes_Own_Delimiters_In_Nested_Text() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();

        var document = MarkdownReader.Parse("++a\\+\\+b++\n\n^x\\^y^\n\n~h\\~i~", options);

        Assert.Collection(
            document.Blocks.Cast<ParagraphBlock>(),
            paragraph => Assert.Equal("++a\\+\\+b++", paragraph.Inlines.RenderMarkdown()),
            paragraph => Assert.Equal("^x\\^y^", paragraph.Inlines.RenderMarkdown()),
            paragraph => Assert.Equal("~h\\~i~", paragraph.Inlines.RenderMarkdown()));
    }

    [Fact]
    public void NoPipe_Table_Body_Does_Not_Terminate_On_Disabled_Heading_Syntax() {
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        options.Headings = false;

        var document = MarkdownReader.Parse("""
            | A |
            |---|
            # value
            """, options);

        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));
        var row = Assert.Single(table.Rows);
        Assert.Equal("# value", Assert.Single(row));
    }
}
