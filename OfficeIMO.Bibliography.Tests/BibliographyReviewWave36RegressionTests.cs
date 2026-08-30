namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave36RegressionTests {
    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Normalized_Bib_keys_remain_unique_and_are_reported(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        document.Items.Add(new BibliographyItem { Key = "a b", Type = BibliographyItemType.Book, Title = "First" });
        document.Items.Add(new BibliographyItem { Key = "a_b", Type = BibliographyItemType.Book, Title = "Second" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyDocument reopened = BibliographyDocument.Parse(written.Content, format).Document;
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Equal(new[] { "a_b", "a_b-2" }, reopened.Items.Select(static item => item.Key));
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV216" && diagnostic.Field == "key");
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV217" && diagnostic.Field == "key");
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex, nameof(BibliographyName.Given), " John  Paul ")]
    [InlineData(BibliographyFormat.BibLatex, nameof(BibliographyName.Given), " John  Paul ")]
    [InlineData(BibliographyFormat.BibTex, nameof(BibliographyName.Family), " Doe ")]
    [InlineData(BibliographyFormat.BibLatex, nameof(BibliographyName.Suffix), " Jr. ")]
    [InlineData(BibliographyFormat.BibTex, nameof(BibliographyName.NonDroppingParticle), " van ")]
    [InlineData(BibliographyFormat.BibLatex, nameof(BibliographyName.DroppingParticle), " de ")]
    public void Bib_component_whitespace_is_diagnosed_before_strict_output(BibliographyFormat format, string property, string value) {
        var name = new BibliographyName { Given = "John", Family = "Doe" };
        switch (property) {
            case nameof(BibliographyName.Given): name.Given = value; break;
            case nameof(BibliographyName.Family): name.Family = value; break;
            case nameof(BibliographyName.Suffix): name.Suffix = value; break;
            case nameof(BibliographyName.NonDroppingParticle): name.NonDroppingParticle = value; break;
            case nameof(BibliographyName.DroppingParticle): name.DroppingParticle = value; break;
        }
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "name", Type = BibliographyItemType.Book };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, name));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV226" && diagnostic.Field == "contributors");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EndNote_detection_accepts_custom_roots_with_direct_same_namespace_records(bool namespaced) {
        string source = namespaced
            ? "<?xml version=\"1.0\"?><library xmlns=\"urn:example\"><metadata/><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type></record></records></library>"
            : "<!--before--><library><metadata/><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type></record></records></library>";

        BibliographyReadResult read = BibliographyDocument.Parse(source);

        Assert.Equal(BibliographyFormat.EndNoteXml, read.Document.SourceFormat);
        Assert.Equal("1", Assert.Single(read.Document.Items).Key);
    }

    [Fact]
    public void EndNote_detection_does_not_promote_nested_records_elements() {
        const string source = "<library><section><records><record/></records></section></library>";

        Assert.Throws<FormatException>(() => BibliographyDocument.Parse(source));
    }

    [Theory]
    [InlineData("contributors")]
    [InlineData("titles")]
    [InlineData("periodical")]
    [InlineData("dates")]
    [InlineData("urls")]
    [InlineData("keywords")]
    public void Whitespace_only_EndNote_containers_survive_unrelated_strict_edits(string container) {
        string source = $"<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><{container}>\n  </{container}></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Publisher = "Publisher";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Contains(reopened.NativeFields, field => field.Format == BibliographyFormat.EndNoteXml && string.Equals(field.Name, container, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void CSL_conversion_reopens_unknown_native_types_without_loss() {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        document.Items.Add(new BibliographyItem { Key = "custom", Type = BibliographyItemType.Unknown, NativeType = "custom-article", Title = "Title" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.CslJson, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        Assert.Equal(BibliographyItemType.Unknown, reopened.Type);
        Assert.Equal("custom-article", reopened.NativeType);
    }
}
