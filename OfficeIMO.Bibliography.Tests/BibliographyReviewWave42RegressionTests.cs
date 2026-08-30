namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave42RegressionTests {
    [Theory]
    [InlineData(17)]
    [InlineData(999)]
    public void Named_unknown_EndNote_types_diagnose_noncanonical_numeric_codes(int code) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Custom Type\">" + code.ToString(System.Globalization.CultureInfo.InvariantCulture) + "</ref-type></record></records></xml>";
        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml);

        BibliographyDiagnostic diagnostic = Assert.Single(read.Diagnostics, diagnostic => diagnostic.Code == "BIBEND004" && diagnostic.Field == "ref-type");
        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            read.Document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(code.ToString(System.Globalization.CultureInfo.InvariantCulture), diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV222" && diagnostic.Field == "ref-type");
    }

    [Fact]
    public void Recognized_RIS_type_tokens_that_require_normalization_block_strict_output() {
        const string source = "TY  - JOUR \nID  - x\nTI  - Before\nER  -\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;
        document.Items[0].Title = "After";

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Equal("JOUR ", document.Items[0].NativeType);
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV200" && diagnostic.Field == "type");
    }

    [Theory]
    [InlineData("element")]
    [InlineData("records-element")]
    public void EndNote_native_entry_names_must_match_their_XML_elements(string kind) {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, kind, "<bar>value</bar>", "foo"));

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV117" && diagnostic.Field == "foo");
    }

    [Fact]
    public void Empty_literal_years_round_trip_through_classic_BibTeX() {
        var document = new BibliographyDocument(BibliographyFormat.BibTex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Literal = string.Empty });
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibTex).Document.Items).Dates);

        Assert.Equal(string.Empty, reopened.Literal);
    }

    [Fact]
    public void Unknown_NBIB_items_cannot_use_unknown_publication_types_as_round_trip_evidence() {
        const string source = "PT  - Custom Type\nTI  - Before\nPMID- 1\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;
        document.Items[0].Type = BibliographyItemType.Unknown;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV200" && diagnostic.Field == "type");
    }
}
