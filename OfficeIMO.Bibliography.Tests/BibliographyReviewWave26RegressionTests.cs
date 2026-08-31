namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave26RegressionTests {
    [Theory]
    [InlineData("author")]
    [InlineData("editor")]
    [InlineData("translator")]
    [InlineData("recipient")]
    [InlineData("interviewer")]
    [InlineData("composer")]
    [InlineData("collection-editor")]
    public void Empty_CSL_contributor_arrays_survive_unrelated_strict_edits(string property) {
        string source = "[{\"id\":\"x\",\"type\":\"book\",\"title\":\"Before\",\"" + property + "\":[]}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        BibliographyNativeField field = Assert.Single(reopened.NativeFields, candidate => candidate.Name == property);
        Assert.Equal("[]", field.RawValue);
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris, "contributors")]
    [InlineData(BibliographyFormat.Ris, "identifiers")]
    [InlineData(BibliographyFormat.Ris, "keywords")]
    [InlineData(BibliographyFormat.Ris, "notes")]
    [InlineData(BibliographyFormat.Nbib, "contributors")]
    [InlineData(BibliographyFormat.Nbib, "identifiers")]
    [InlineData(BibliographyFormat.Nbib, "keywords")]
    [InlineData(BibliographyFormat.Nbib, "notes")]
    public void Tagged_collection_serialization_observes_cancellation(BibliographyFormat format, string collection) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "x", Type = format == BibliographyFormat.Nbib ? BibliographyItemType.ArticleJournal : BibliographyItemType.Book };
        for (int index = 0; index < 200_000; index++) {
            if (collection == "contributors") item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Family" + index }));
            else if (collection == "identifiers") item.Identifiers.Add(new BibliographyIdentifier("custom", index.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            else if (collection == "keywords") item.Keywords.Add("keyword" + index);
            else item.Notes.Add("note" + index);
        }
        if (format == BibliographyFormat.Nbib) item.Identifiers.Insert(0, new BibliographyIdentifier("PMID", "x"));
        document.Items.Add(item);
        BibliographyCancellationTest.AssertObserved(token => {
            if (format == BibliographyFormat.Ris) TaggedCodec.WriteRis(document, new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, new BibliographyConversionReport(), token);
            else TaggedCodec.WriteNbib(document, new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, new BibliographyConversionReport(), token);
        });
    }

    [Theory]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Null_valued_empty_dates_are_diagnosed_before_strict_output(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = format == BibliographyFormat.Nbib ? BibliographyItemType.ArticleJournal : BibliographyItemType.Book };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued });
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV242" && diagnostic.Field == "dates.Issued");
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Literal_names_ending_in_commas_reopen_as_literals(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Literal = "Acme," }));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyName reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items[0].Contributors).Name;

        Assert.Equal("Acme,", reopened.Literal);
        Assert.Null(reopened.Family);
    }

    [Fact]
    public void EndNote_preserves_distinct_nonnumeric_year_and_publication_date_components() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles><dates><year>n.d.</year><pub-dates><date>Spring</date></pub-dates></dates></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        BibliographyDate date = Assert.Single(reopened.Dates);
        Assert.Equal("Spring", date.Literal);
        Assert.Contains(date.NativeFields, field => field.Format == BibliographyFormat.EndNoteXml && field.Name == "year" && field.Value == "n.d.");
    }
}
