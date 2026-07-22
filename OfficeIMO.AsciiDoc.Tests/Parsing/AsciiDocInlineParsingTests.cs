namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocInlineParsingTests {
    [Fact]
    public void CommonInlineForms_AreTypedNestedAndLossless() {
        const string source =
            "A *strong _nested_* and Sara**h** with `code`, \\*literal*, {product}, " +
            "<<intro,Introduction>>, [[local]], image:icon.svg[Icon], stem:[x^2], and +{raw}+.\r\n";

        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocParagraph paragraph = Assert.Single(document.BlocksOfType<AsciiDocParagraph>());

        AsciiDocFormattedInline[] formatted = paragraph.Inlines.Items.OfType<AsciiDocFormattedInline>().ToArray();
        Assert.Contains(formatted, item => item.Style == AsciiDocInlineStyle.Strong && !item.IsUnconstrained);
        Assert.Contains(formatted, item => item.Style == AsciiDocInlineStyle.Strong && item.IsUnconstrained);
        Assert.Contains(formatted, item => item.Style == AsciiDocInlineStyle.Monospace);
        AsciiDocFormattedInline strong = formatted.First(item => !item.IsUnconstrained && item.Style == AsciiDocInlineStyle.Strong);
        Assert.Contains(strong.Content.Items, item => item is AsciiDocFormattedInline nested && nested.Style == AsciiDocInlineStyle.Emphasis);
        Assert.Single(paragraph.Inlines.Items.OfType<AsciiDocAttributeReferenceInline>());
        Assert.Single(paragraph.Inlines.Items.OfType<AsciiDocCrossReferenceInline>());
        Assert.Single(paragraph.Inlines.Items.OfType<AsciiDocAnchorInline>());
        Assert.Single(paragraph.Inlines.Items.OfType<AsciiDocMacroInline>(), item => item is not AsciiDocStemInline);
        Assert.Single(paragraph.Inlines.Items.OfType<AsciiDocStemInline>());
        Assert.Single(paragraph.Inlines.Items.OfType<AsciiDocPassthroughInline>());
        Assert.Equal(source, document.ToAsciiDoc());
        Assert.True(document.SyntaxTree.IsLossless);
    }

    [Fact]
    public void InlineEdits_RegenerateOnlyChangedNodes() {
        const string source = "Before *old* {name} <<old-id,Old label>> after.\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocParagraph paragraph = Assert.Single(document.BlocksOfType<AsciiDocParagraph>());

        AsciiDocFormattedInline strong = Assert.Single(paragraph.Inlines.Items.OfType<AsciiDocFormattedInline>());
        Assert.Single(strong.Content.Items.OfType<AsciiDocTextInline>()).Text = "new";
        Assert.Single(paragraph.Inlines.Items.OfType<AsciiDocAttributeReferenceInline>()).Name = "product";
        AsciiDocCrossReferenceInline xref = Assert.Single(paragraph.Inlines.Items.OfType<AsciiDocCrossReferenceInline>());
        xref.Target = "new-id";
        xref.Text = "New label";

        Assert.True(document.IsModified);
        Assert.Equal("Before *new* {product} <<new-id,New label>> after.\n", document.ToAsciiDoc());
    }

    [Fact]
    public void PunctuationAndEscapes_DoNotCreateFalseFormatting() {
        const string source = "2 * 3, snake_case, and \\*literal*.\n";

        AsciiDocParagraph paragraph = Assert.Single(AsciiDocDocument.Parse(source).Document.BlocksOfType<AsciiDocParagraph>());

        Assert.Empty(paragraph.Inlines.Items.OfType<AsciiDocFormattedInline>());
        Assert.Equal(source.Substring(0, source.Length - 1), paragraph.Inlines.ToAsciiDoc());
    }

    [Fact]
    public void MultilineInlineSequence_RetainsInternalMixedLineEndings() {
        const string source = "First *bold* line\r\nsecond {name} line\rLast line\n";

        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocParagraph paragraph = Assert.Single(document.BlocksOfType<AsciiDocParagraph>());

        Assert.Equal("First *bold* line\r\nsecond {name} line\rLast line", paragraph.Inlines.Syntax.OriginalText);
        Assert.Equal(source, document.ToAsciiDoc());
    }

    [Fact]
    public void InlineNodeLimit_IsEnforced() {
        var options = new AsciiDocParseOptions { MaximumInlineNodeCount = 2 };

        Assert.Throws<InvalidDataException>(() => AsciiDocDocument.Parse("*one* {two} three", options));
    }

    [Fact]
    public void LongEscapeRuns_AreParsedLosslesslyWithoutBackwardRescanning() {
        string source = new string('\\', 100_000) + "*literal*\n";

        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;

        Assert.Equal(source, document.ToAsciiDoc());
    }
}
