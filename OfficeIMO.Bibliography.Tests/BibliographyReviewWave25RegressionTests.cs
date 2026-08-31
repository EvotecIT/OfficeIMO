namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave25RegressionTests {
    [Fact]
    public void Empty_primary_EndNote_URL_survives_an_unrelated_strict_edit() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles><urls><related-urls><url/></related-urls></urls></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal(string.Empty, reopened.Url);
    }

    [Theory]
    [InlineData("DO")]
    [InlineData("SN")]
    [InlineData("AN")]
    public void Blank_RIS_identifier_tags_remain_native_after_other_edits(string tag) {
        string source = "TY  - BOOK\nID  - x\n" + tag + "  - \nTI  - Before\nER  -\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.Contains(reopened.NativeFields, field => field.Name == tag && field.Value.Length == 0);
    }

    [Fact]
    public void NBIB_identifier_scheme_casing_survives_canonical_edits() {
        const string source = "PMID- 1\nAID - value [CustomID]\nTI  - Before\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Nbib).Document.Items);

        Assert.Contains(reopened.Identifiers, identifier => identifier.Scheme == "CustomID" && identifier.Value == "value");
    }

    [Fact]
    public void NBIB_compact_initials_preserve_supplementary_Unicode_scalars() {
        string scalar = char.ConvertFromUtf32(0x10437);
        var document = new BibliographyDocument(BibliographyFormat.Nbib);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.ArticleJournal, Title = "Unicode" };
        item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Smith", Given = scalar + " Name" }));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyContributor reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Nbib).Document.Items[0].Contributors);

        Assert.Contains("AU  - Smith " + scalar + "N", written.Content, StringComparison.Ordinal);
        Assert.Equal(scalar + " Name", reopened.Name.Given);
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Family_only_multiword_names_reopen_as_family_only(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = format == BibliographyFormat.Nbib ? BibliographyItemType.ArticleJournal : BibliographyItemType.Book, Title = "Names" };
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Van Helsing" }));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyContributor reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items[0].Contributors);

        Assert.Equal("Van Helsing", reopened.Name.Family);
        Assert.True(string.IsNullOrEmpty(reopened.Name.Given));
        Assert.Null(reopened.Name.Literal);
    }
}
