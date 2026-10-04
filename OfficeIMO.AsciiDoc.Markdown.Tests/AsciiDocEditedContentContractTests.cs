namespace OfficeIMO.AsciiDoc.Markdown.Tests;

public sealed class AsciiDocEditedContentContractTests {
    [Fact]
    public void ScalarAndInlineEditsAgreeAcrossNativeMarkdownAndWordProjection() {
        const string source = "= Original title\n\nOriginal paragraph\n\n* Original item\n\nNOTE: Original note\n\nOriginal term:: Original definition\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source);
        document.BlocksOfType<AsciiDocHeading>().Single().Title = "Changed title";
        AsciiDocParagraph paragraph = document.BlocksOfType<AsciiDocParagraph>().Single();
        paragraph.Text = "Changed *paragraph*";
        paragraph.Inlines.Items.OfType<AsciiDocTextInline>().Single().Text = "Latest ";
        document.BlocksOfType<AsciiDocListBlock>().Single().Items[0].Text = "Changed item";
        document.BlocksOfType<AsciiDocAdmonitionBlock>().Single().Text = "Changed note";
        AsciiDocDescriptionListItem definition = document.BlocksOfType<AsciiDocDescriptionListBlock>().Single().Items[0];
        definition.Term = "Changed term";
        definition.Description = "Changed definition";

        string native = document.ToAsciiDoc();
        AsciiDocToMarkdownResult conversion = document.ToMarkdownDocumentResult();
        string markdown = conversion.Value.ToMarkdown();
        foreach (string value in new[] { "Changed title", "Latest", "Changed item", "Changed note", "Changed term", "Changed definition" }) {
            Assert.Contains(value, native, StringComparison.Ordinal);
            Assert.Contains(value, markdown, StringComparison.Ordinal);
        }
        Assert.Equal("Latest *paragraph*", paragraph.Text);
        Assert.DoesNotContain("Original", markdown, StringComparison.Ordinal);
        using var word = conversion.Value.ToWordDocument();
        Assert.Contains("Latest", string.Join(" ", word.Paragraphs.Select(item => item.Text)), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("====")]
    [InlineData("****")]
    [InlineData("--")]
    public void CompoundContainerPreservesAllParagraphsAndLists(string delimiter) {
        string source = delimiter + "\nFirst paragraph.\n\nSecond paragraph.\n\n* List item\n" + delimiter + "\n";
        AsciiDocToMarkdownResult result = AsciiDocDocument.Parse(source).ToMarkdownDocumentResult();
        Assert.Equal(2, result.Value.Blocks.OfType<ParagraphBlock>().Count());
        Assert.Single(result.Value.Blocks.OfType<UnorderedListBlock>());
        string markdown = result.Value.ToMarkdown();
        Assert.Contains("First paragraph.", markdown, StringComparison.Ordinal);
        Assert.Contains("Second paragraph.", markdown, StringComparison.Ordinal);
        Assert.Contains("List item", markdown, StringComparison.Ordinal);
        Assert.Equal(AsciiDocMarkdownConversionOutcome.Simplified, Assert.Single(result.Report.Diagnostics).Outcome);
    }

    [Fact]
    public void AttributeAssignmentsCaptureDocumentOrderAndUnsetFinalMetadata() {
        const string source = ":version: one\n:captured: {version}\n:draft:\n\n{version}\n\n:version: two\n:draft!:\n\n{version}\n\n{captured}\n";
        AsciiDocToMarkdownResult result = AsciiDocDocument.Parse(source).ToMarkdownDocumentResult();
        string[] paragraphs = result.Value.Blocks.OfType<ParagraphBlock>().Select(item => MarkdownDoc.Create().Add(item).ToMarkdown()).ToArray();
        Assert.Equal(new[] { "one", "two", "one" }, paragraphs.Select(value => value.Trim()));
        Assert.DoesNotContain("draft:", result.Value.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void CompoundAttributesAndEditsRemainEffectiveAfterTheContainer() {
        AsciiDocDocument document = AsciiDocDocument.Parse(":value: before\n\n====\n{value}\n\n:value: inside\n\n{value}\n====\n\n{value}\n");
        AsciiDocDelimitedBlock container = document.BlocksOfType<AsciiDocDelimitedBlock>().Single();
        container.Body!.BlocksOfType<AsciiDocParagraph>().Last().Text = "Edited {value}";
        var paragraphs = document.ToMarkdownDocumentResult().Value.Blocks.OfType<ParagraphBlock>()
            .Select(block => MarkdownDoc.Create().Add(block).ToMarkdown().Trim()).ToArray();
        Assert.Equal(new[] { "before", "Edited inside", "inside" }, paragraphs);
    }

    [Fact]
    public void CompoundFallbackDiagnosticsUseOriginalSourceLines() {
        AsciiDocDocument document = AsciiDocDocument.Parse("= Title\n\n====\ncustom:value[]\n====\n");
        AsciiDocMarkdownConversionDiagnostic inline = Assert.Single(document.ToMarkdownDocumentResult().Report.Diagnostics, diagnostic => diagnostic.Code == "ADOCMD102");
        Assert.Equal(4, inline.SourceSpan.Start.Line);
    }
}
