namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave28RegressionTests {
    [Theory]
    [InlineData("contributors")]
    [InlineData("identifiers")]
    [InlineData("keywords")]
    [InlineData("notes")]
    [InlineData("native-fields")]
    [InlineData("dates")]
    [InlineData("document-native-entries")]
    public void EndNote_collection_serialization_observes_cancellation(string collection) {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        var contributor = new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Family" });
        var identifier = new BibliographyIdentifier("DOI", "10.1000/example");
        var nativeField = new BibliographyNativeField(BibliographyFormat.BibTex, "custom", "value");
        var nativeEntry = new BibliographyNativeEntry(BibliographyFormat.BibTex, "comment", "value");
        var date = new BibliographyDate { Role = BibliographyDateRole.Accessed, Year = 2026 };
        for (int index = 0; index < 200_000; index++) {
            if (collection == "contributors") item.Contributors.Add(contributor);
            else if (collection == "identifiers") item.Identifiers.Add(identifier);
            else if (collection == "keywords") item.Keywords.Add("keyword");
            else if (collection == "notes") item.Notes.Add("note");
            else if (collection == "native-fields") item.NativeFields.Add(nativeField);
            else if (collection == "dates") item.Dates.Add(date);
            else document.NativeEntries.Add(nativeEntry);
        }
        document.Items.Add(item);
        using var cancellation = new CancellationTokenSource();
        cancellation.CancelAfter(1);

        Assert.Throws<OperationCanceledException>(() =>
            EndNoteXmlCodec.Write(document, new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, new BibliographyConversionReport(), cancellation.Token));
    }

    [Theory]
    [InlineData(BibliographyFormat.Nbib, "publisher")]
    [InlineData(BibliographyFormat.Nbib, "publisher-place")]
    [InlineData(BibliographyFormat.Nbib, "edition")]
    [InlineData(BibliographyFormat.Nbib, "URL")]
    [InlineData(BibliographyFormat.Nbib, "collection-title")]
    [InlineData(BibliographyFormat.Ris, "collection-title")]
    public void Empty_unsupported_tagged_properties_are_diagnosed_before_strict_output(BibliographyFormat format, string field) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = format == BibliographyFormat.Nbib ? BibliographyItemType.ArticleJournal : BibliographyItemType.Book };
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        switch (field) {
            case "publisher": item.Publisher = string.Empty; break;
            case "publisher-place": item.PublisherPlace = string.Empty; break;
            case "edition": item.Edition = string.Empty; break;
            case "URL": item.Url = string.Empty; break;
            case "collection-title": item.CollectionTitle = string.Empty; break;
        }
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV203" && diagnostic.Field == field);
    }

    [Theory]
    [InlineData("contributors", "<contributors><author>Jane Doe</author></contributors>", "Jane Doe")]
    [InlineData("contributors", "<contributors><reviewers><author>Jane Doe</author></reviewers></contributors>", "Jane Doe")]
    [InlineData("dates", "<dates><date>2026</date></dates>", "2026")]
    [InlineData("urls", "<urls><url>https://example.com</url></urls>", "https://example.com")]
    [InlineData("keywords", "<keywords><wrapper><keyword>alpha</keyword></wrapper></keywords>", "alpha")]
    public void Unsupported_EndNote_container_shapes_remain_native(string container, string unsupportedXml, string marker) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles>" + unsupportedXml + "</record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        BibliographyItem item = Assert.Single(document.Items);
        AssertUnsupportedOwnerIsEmpty(item, container);
        Assert.Contains(item.NativeFields, field => field.Format == BibliographyFormat.EndNoteXml && field.Name == container && field.RawValue!.Contains(marker, StringComparison.Ordinal));
        item.Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        AssertUnsupportedOwnerIsEmpty(reopened, container);
        Assert.Contains(reopened.NativeFields, field => field.Name == container && field.RawValue!.Contains(marker, StringComparison.Ordinal));
    }

    private static void AssertUnsupportedOwnerIsEmpty(BibliographyItem item, string container) {
        if (container == "contributors") Assert.Empty(item.Contributors);
        else if (container == "dates") Assert.Empty(item.Dates);
        else if (container == "urls") Assert.Null(item.Url);
        else if (container == "keywords") Assert.Empty(item.Keywords);
    }
}
