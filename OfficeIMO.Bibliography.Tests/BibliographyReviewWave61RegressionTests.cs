namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave61RegressionTests {
    [Theory]
    [InlineData(null)]
    [InlineData("https://primary.example")]
    public void EndNote_interleaved_additional_URLs_keep_native_field_order(string? primaryUrl) {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Title", Url = primaryUrl };
        item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, "url", "https://one.example"));
        item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, "custom", "middle"));
        item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, "url", "https://two.example"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal(primaryUrl, reopened.Url);
        Assert.Equal(new[] { "url", "custom", "url" }, reopened.NativeFields.Select(static field => field.Name));
        Assert.Equal(new[] { "https://one.example", "middle", "https://two.example" }, reopened.NativeFields.Select(static field => field.Value));
    }

    [Fact]
    public void EndNote_additional_URLs_remain_native_when_the_primary_URL_is_cleared() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><urls><related-urls><url>https://primary.example</url></related-urls></urls><urls><related-urls><url>https://additional.example</url></related-urls></urls></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Url = null;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Null(reopened.Url);
        BibliographyNativeField additional = Assert.Single(reopened.NativeFields, static field => field.Name == "url");
        Assert.Equal("https://additional.example", additional.Value);
    }

    [Fact]
    public void NBIB_initials_keep_component_semantics_without_split_allocation() {
        Assert.Equal("JLM", TaggedCodec.Initials("Jean-Luc María"));
    }

    [Fact]
    public void NBIB_initial_derivation_observes_cancellation_during_large_names() {
        string given = string.Concat(Enumerable.Repeat("A-", 8 * 1024 * 1024));

        BibliographyCancellationTest.AssertObserved(token => TaggedCodec.Initials(given, token));
    }
}
