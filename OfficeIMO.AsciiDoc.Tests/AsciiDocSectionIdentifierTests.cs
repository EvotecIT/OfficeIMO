namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocSectionIdentifierTests {
    // Expected IDs were checked with Asciidoctor 2.0.26, independently of this implementation.
    [Theory]
    [InlineData("Wiley & Sons, Inc.", "_wiley_sons_inc")]
    [InlineData("A *Bold* _Title_", "_a_bold_title")]
    [InlineData("Łódź Ελληνικά 日本語", "_łódź_ελληνικά_日本語")]
    [InlineData("A &copy; B &#169; C", "_a_b_c")]
    [InlineData("Read https://example.test[The Guide]", "_read_the_guide")]
    [InlineData("A <b>Literal</b> Title", "_a_bliteralb_title")]
    [InlineData("A--B", "_ab")]
    [InlineData("A...B", "_ab")]
    [InlineData("A....B", "_a_b")]
    [InlineData("A---B", "_a_b")]
    public void VisibleTitleSemanticsGenerateInteroperableIds(string title, string expected) {
        AsciiDocDocument document = AsciiDocDocument.Parse("== " + title + "\n");
        AsciiDocReferenceCatalog catalog = AsciiDocReferenceCatalog.Create(document);
        var target = Assert.Single(catalog.Targets).Value;
        Assert.Equal(expected, target.Id);
        Assert.True(target.IsGenerated);
        Assert.Equal(expected, catalog.GetBlockId(document.BlocksOfType<AsciiDocHeading>().Single()));
        Assert.Empty(catalog.Diagnostics);
        Assert.Equal("== " + title + "\n", document.ToAsciiDoc());
    }

    [Theory]
    [InlineData(":idprefix: topic-\n:idseparator: -\n", "A.B -- C", "topic-a-bc", "topic-a-bc-2")]
    [InlineData(":idseparator:\n", "A.B -- C", "_a.bc", "_a.bc2")]
    [InlineData(":idprefix:\n", ". Spaces .", "spaces", "spaces_2")]
    [InlineData(":idseparator: xy\n", "A B", "_axb", "_axbx2")]
    public void PrefixAndSeparatorControlDuplicateSuffixes(string attributes, string title, string first, string second) {
        AsciiDocDocument document = AsciiDocDocument.Parse(attributes + "\n== " + title + "\n\n== " + title + "\n");
        AsciiDocReferenceCatalog catalog = AsciiDocReferenceCatalog.Create(document);
        Assert.Equal(new[] { first, second }, document.BlocksOfType<AsciiDocHeading>().Select(catalog.GetBlockId));
    }

    [Fact]
    public void SourceOrderAttributesExplicitAnchorsAndDisabledSectionsShareOneNamespace() {
        const string source = "[[_same]]\nA paragraph.\n\n:project: First\n\n== {project}\n\n:project: Second\n\n== {project}\n\n:!sectids:\n\n== Disabled\n\n[[custom,Label]]\n== Explicit\n\n:sectids:\n\n== Same\n\n== Same\n\n== Inline [[inline-id]]\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source);
        AsciiDocReferenceCatalog catalog = AsciiDocReferenceCatalog.Create(document);
        Assert.Equal(new string?[] { "_first", "_second", null, "custom", "_same_2", "_same_3", "inline-id" }, document.BlocksOfType<AsciiDocHeading>().Select(catalog.GetBlockId));
        Assert.Equal("Label", catalog.Targets["custom"].Label);
        Assert.False(catalog.Targets["inline-id"].IsGenerated);
        Assert.Empty(catalog.Diagnostics);
        document.BlocksOfType<AsciiDocHeading>().First().Title = "Edited";
        Assert.Equal("_edited", AsciiDocReferenceCatalog.Create(document).GetBlockId(document.BlocksOfType<AsciiDocHeading>().First()));
        Assert.Equal("_first", catalog.GetBlockId(document.BlocksOfType<AsciiDocHeading>().First()));
    }

    [Fact]
    public void DocumentTitlesAndEmptySectionIdsHaveNoGeneratedTarget() {
        AsciiDocDocument document = AsciiDocDocument.Parse("= Document\n\n== !!!\n\n== Section\n");
        AsciiDocReferenceCatalog catalog = AsciiDocReferenceCatalog.Create(document);
        Assert.Equal(new string?[] { null, null, "_section" }, document.BlocksOfType<AsciiDocHeading>().Select(catalog.GetBlockId));
        Assert.Equal("ADOCREF006", Assert.Single(catalog.Diagnostics).Code);
    }

    [Fact]
    public void PassthroughTitlesAndSubstitutionOverridesRequireExplicitInteropReview() {
        AsciiDocDocument document = AsciiDocDocument.Parse("== A pass:[<b>Rich</b>] Title\n\n[subs=\"-replacements\"]\n== A--B\n");
        AsciiDocReferenceCatalog catalog = AsciiDocReferenceCatalog.Create(document);
        Assert.Equal("A Rich Title", catalog.Targets["_a_rich_title"].Label);
        Assert.Equal(2, catalog.Diagnostics.Count(diagnostic => diagnostic.Code == "ADOCREF005"));
    }
}
