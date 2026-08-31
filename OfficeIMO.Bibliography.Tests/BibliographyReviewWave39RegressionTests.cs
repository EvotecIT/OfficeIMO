namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave39RegressionTests {
    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Non_CSL_writers_diagnose_given_only_structured_names(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "123", Type = BibliographyItemType.Book, Title = "Names" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Given = "John" }));
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "123"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV244" && diagnostic.Field == "contributors");
    }

    [Fact]
    public void Unknown_EndNote_native_types_survive_strict_same_format_edits() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Custom Type\">13</ref-type><titles><title>Before</title></titles></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal(BibliographyItemType.Unknown, reopened.Type);
        Assert.Equal("Custom Type", reopened.NativeType);
    }

    [Fact]
    public void EndNote_custom_root_detection_enforces_the_configured_nesting_limit() {
        const string source = "<library><extension><one><two/></one></extension><records/></library>";
        var options = new BibliographyReadOptions { MaximumNestingDepth = 2 };

        Assert.Throws<InvalidDataException>(() => BibliographyDocument.Parse(source, options));
    }

    [Theory]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Syntactically_valid_non_calendar_dates_round_trip_strictly(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "123", Type = BibliographyItemType.Book, Title = "Dates" };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Year = 2023, Month = 2, Day = 31 });
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "123"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items).Dates);

        Assert.Equal(2023, reopened.Year);
        Assert.Equal(2, reopened.Month);
        Assert.Equal(31, reopened.Day);
    }

    [Theory]
    [InlineData("{\"a\": 1}")]
    [InlineData("[1, 2]")]
    public void Edited_CSL_aggregate_native_values_preserve_their_formatting(string editedValue) {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"x\",\"type\":\"book\",\"custom\":{\"before\":true}}]", BibliographyFormat.CslJson).Document;
        BibliographyNativeField field = Assert.Single(document.Items[0].NativeFields);
        field.Value = editedValue;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyNativeField reopened = Assert.Single(Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items).NativeFields);

        Assert.Equal(editedValue, reopened.Value);
    }

    [Fact]
    public void Duplicate_EndNote_root_attribute_carriers_are_diagnosed_before_omission() {
        const string source = "<xml first=\"1\"><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, "attributes", "<xml second=\"2\" />", "xml"));

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV248" && diagnostic.Field == "xml");
    }
}
