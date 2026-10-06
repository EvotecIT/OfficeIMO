using System;
using System.Collections.Generic;
using System.Linq;
using global::ChartForgeX.Diagnostics;
using global::ChartForgeX.Markup;
using OfficeIMO.ChartForgeX.Markdown;
using OfficeIMO.Markdown;
using OfficeIMO.Markup;
using OfficeIMO.Markup.PowerPoint;
using OfficeIMO.Markup.Word;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.ChartForgeX.Markdown.Tests;

public sealed class MermaidMarkupRegressionTests {
    private const string Fence = "```mermaid title=\"Approval\"\nflowchart LR\nA --> B\n```";

    [Theory]
    [InlineData("# Report\n\nIntro\n\n", "", 6)]
    [InlineData("---\nprofile: document\n---\n\n# Report\n\nIntro\n\n", "", 10)]
    [InlineData("@slide title=\"Report\"\n\nIntro\n\n", "", 4)]
    [InlineData("~~~~officeimo-slide\ntitle=\"Report\"\n\nIntro\n\n", "\n~~~~", 4)]
    public void MarkupTransformDiagnosticsRetainDocumentOrFragmentLocations(string prefix, string suffix, int line) {
        var errors = new List<MarkupDiagnostic>();
        var reader = MarkdownReaderOptions.CreateOfficeIMOProfile();
        reader.DocumentTransforms.Add(MermaidMarkdownAdapter.CreateTransform(errors.Add));
        OfficeMarkupParser.Parse(prefix + "```mermaid\nnot-a-diagram\n```" + suffix,
            new OfficeMarkupParserOptions { MarkdownOptions = reader });
        Assert.Equal(line, Assert.Single(errors, item => item.Severity == VisualDiagnosticSeverity.Error).Line);
    }

    [Fact]
    public void WordMarkupPreservesNestedListImagesAndDoesNotRepeatTheLeadText() {
        string first = "- Approval\n\n" + Indent(Fence + "\n_Approval caption_", 2);
        string source = first + "\n\n  - Detail\n\n" + Indent(Fence, 4);
        var parsed = Parse(source, OfficeMarkupProfile.Document);
        Assert.Equal(2, parsed.Document.DescendantsAndSelf().OfType<OfficeMarkupImageBlock>().Count());
        using var word = parsed.Document.ToWordDocument();
        Assert.Equal(2, word.Images.Count);
        Assert.Single(word.Paragraphs, paragraph => paragraph.Text == "Approval");
        Assert.Single(word.Paragraphs, paragraph => paragraph.Text == "Detail");
        Assert.Contains(word.Paragraphs, paragraph => paragraph.Text == "Approval caption");
    }

    [Theory]
    [InlineData("blank")]
    [InlineData("content")]
    [InlineData("process")]
    [InlineData("timeline")]
    public void PowerPointLayoutsPreserveRichListContent(string layout) {
        string source = "@slide title=\"Report\" layout=" + layout + "\n\n- Approval\n\n"
            + Indent(Fence + "\n_Approval caption_", 2) + "\n\n- Complete";
        var parsed = Parse(source, OfficeMarkupProfile.Presentation);
        using var deck = parsed.Document.ToPowerPointPresentation();
        var slide = Assert.Single(deck.Slides);
        Assert.Single(slide.Shapes.OfType<PowerPointPicture>());
        Assert.Contains(slide.TextBoxes, box => box.Text == "Approval caption");
        Assert.Contains(slide.TextBoxes, box => box.Text.Contains("Complete"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ColumnBodiesUseTheCallerTransformAndPreserveRichLists(bool nestedList) {
        string body = nestedList ? "- Approval\n\n" + Indent(Fence, 2) : Fence;
        string source = "@slide title=\"Comparison\" layout=comparison\n\n::columns gap=4%\n\n::left width=48%\n"
            + body + "\n\n::right width=48%\n## Complete\n- Summary";
        var parsed = Parse(source, OfficeMarkupProfile.Presentation);
        Assert.Single(parsed.Document.DescendantsAndSelf().OfType<OfficeMarkupImageBlock>());
        using var deck = parsed.Document.ToPowerPointPresentation();
        Assert.Single(deck.Slides.SelectMany(slide => slide.Shapes).OfType<PowerPointPicture>());
        Assert.Contains(deck.Slides.SelectMany(slide => slide.TextBoxes), box => box.Text.Contains("Complete"));
    }

    private static OfficeMarkupParseResult Parse(string source, OfficeMarkupProfile profile) {
        var reader = MarkdownReaderOptions.CreateOfficeIMOProfile();
        reader.DocumentTransforms.Add(MermaidMarkdownAdapter.CreateTransform());
        return OfficeMarkupParser.Parse(source, new OfficeMarkupParserOptions { Profile = profile, MarkdownOptions = reader });
    }

    private static string Indent(string source, int width) =>
        string.Join("\n", source.Split('\n').Select(line => new string(' ', width) + line));
}
