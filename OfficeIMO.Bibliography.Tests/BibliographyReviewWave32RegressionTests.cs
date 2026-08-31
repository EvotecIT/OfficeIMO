namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave32RegressionTests {
    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Empty_structured_name_components_are_diagnosed_before_strict_output(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Given = string.Empty, Family = "Doe" }));
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV244" && diagnostic.Field == "contributors");
    }

    [Fact]
    public void CSL_preserves_empty_structured_name_components() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Given = string.Empty, Family = "Doe" }));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyName reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items[0].Contributors).Name;

        Assert.Equal(string.Empty, reopened.Given);
        Assert.Equal("Doe", reopened.Family);
    }

    [Fact]
    public void Lowercase_CSL_identifier_schemes_are_diagnosed_before_strict_output() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Identifiers.Add(new BibliographyIdentifier("doi", "10.1000/example"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV245" && diagnostic.Field == "identifiers.doi");
    }

    [Theory]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Literal_date_trailing_whitespace_round_trips_strictly(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Literal = "n.d. " });
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items[0].Dates);

        Assert.Equal("n.d. ", reopened.Literal);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV246");
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    public void Tagged_whitespace_only_literal_dates_are_diagnosed_when_the_destination_trims_them(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Literal = " " });
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV246" && diagnostic.Field == "dates.Issued.literal");
    }

    [Fact]
    public void CSL_preserves_literal_date_surrounding_whitespace() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Literal = " n.d. " });
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items[0].Dates);

        Assert.Equal(" n.d. ", reopened.Literal);
    }

    [Theory]
    [InlineData("xml")]
    [InlineData("records")]
    [InlineData("record")]
    public void EndNote_structural_mixed_text_is_diagnosed_as_recovered_source_loss(string elementName) {
        string source = CreateMixedTextSource(elementName, "orphan");
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;

        BibliographyDiagnostic diagnostic = Assert.Single(document.Diagnostics, candidate => candidate.Code == "BIBEND005" && candidate.Field == elementName);
        Assert.True(diagnostic.Offset >= 0);
        document.Items[0].Title = "After";

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        Assert.Contains(exception.Report.Diagnostics, candidate => candidate.Code == "BIBCONV222" && candidate.Field == elementName);
    }

    [Theory]
    [InlineData("records")]
    [InlineData("record")]
    public void EndNote_structural_mixed_text_is_bounded_before_DOM_materialization(string elementName) {
        string source = CreateMixedTextSource(elementName, "oversized");

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml, new BibliographyReadOptions { MaximumValueLength = 4 });

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001" && diagnostic.Message.Contains("value length", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("records")]
    [InlineData("record")]
    public void EndNote_structural_mixed_text_counts_as_a_materialized_value(string elementName) {
        const string clean = "<records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles></record></records>";
        var options = new BibliographyReadOptions { MaximumValueCount = 4 };

        BibliographyReadResult baseline = BibliographyDocument.Parse(clean, BibliographyFormat.EndNoteXml, options);
        BibliographyReadResult mixed = BibliographyDocument.Parse(CreateMixedTextSource(elementName, "orphan"), BibliographyFormat.EndNoteXml, options);

        Assert.DoesNotContain(baseline.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
        Assert.Contains(mixed.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001" && diagnostic.Message.Contains("value count", StringComparison.Ordinal));
    }

    private static string CreateMixedTextSource(string elementName, string text) {
        const string record = "<record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles></record>";
        switch (elementName) {
            case "xml": return "<xml>" + text + "<records>" + record + "</records></xml>";
            case "records": return "<records>" + text + record + "</records>";
            case "record": return "<records><record>" + text + "<rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles></record></records>";
            default: throw new ArgumentOutOfRangeException(nameof(elementName));
        }
    }
}
