namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocMacroContentTests {
    [Theory]
    [InlineData("Text footnote:[Don't forget.].", "Don't forget.")]
    [InlineData("Text footnote:[He said \"hello.].", "He said \"hello.")]
    [InlineData("Text footnote:[Use \\] here.].", "Use \\] here.")]
    [InlineData("doublefootnote:[A note.]", "A note.")]
    public void ProseFootnotesAreTypedAndCatalogued(string source, string content) {
        AsciiDocDocument document = AsciiDocDocument.Parse(source);
        AsciiDocParagraph paragraph = Assert.Single(document.BlocksOfType<AsciiDocParagraph>());
        AsciiDocFootnoteInline note = Assert.Single(paragraph.Inlines.Items.OfType<AsciiDocFootnoteInline>());

        Assert.Equal(content, note.AttributeList);
        Assert.Same(note, Assert.Single(AsciiDocReferenceCatalog.Create(document).Footnotes).Value);
        Assert.Equal(source, document.ToAsciiDoc());
        if (source.StartsWith("double", StringComparison.Ordinal))
            Assert.Equal("double", Assert.IsType<AsciiDocTextInline>(paragraph.Inlines.Items[0]).Text);
    }

    [Theory]
    [InlineData("link:guide.pdf[Don't miss it]", "Don't miss it")]
    [InlineData("link:guide.pdf[Say \"hello]", "Say \"hello")]
    [InlineData("link:guide.pdf['Tis a guide]", "'Tis a guide")]
    [InlineData("link:guide.pdf[Don't miss it, title=\"A guide\"]", "Don't miss it, title=\"A guide\"")]
    [InlineData("link:guide.pdf[Read, title=\"Don't miss it\"]", "Read, title=\"Don't miss it\"")]
    [InlineData("image:icon.svg[\"A ] symbol\",title=\"Don't miss it\"]", "\"A ] symbol\",title=\"Don't miss it\"")]
    public void MacroBodiesDistinguishProseQuotesFromQuotedAttributes(string source, string attributes) {
        AsciiDocDocument document = AsciiDocDocument.Parse(source);
        AsciiDocMacroInline macro = Assert.Single(Assert.Single(document.BlocksOfType<AsciiDocParagraph>()).Inlines.Items.OfType<AsciiDocMacroInline>());

        Assert.Equal(attributes, macro.AttributeList);
        Assert.Equal(source, document.ToAsciiDoc());
        if (source.StartsWith("image:", StringComparison.Ordinal)) {
            Assert.Equal("A ] symbol", macro.Attributes.Style);
            Assert.Equal("Don't miss it", macro.Attributes.GetNamedValue("title"));
        }
        if (attributes.StartsWith("Don't", StringComparison.Ordinal))
            Assert.Equal("Don't miss it", macro.Attributes.Style);
    }
}
