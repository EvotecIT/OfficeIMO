namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave50RegressionTests {
    [Theory]
    [InlineData("month")]
    [InlineData("day")]
    [InlineData("end-month")]
    [InlineData("end-day")]
    public void CSL_out_of_range_months_and_days_are_diagnosed_before_strict_output(string component) {
        var date = new BibliographyDate { Role = BibliographyDateRole.Issued, Year = 2024, Month = 1, Day = 2 };
        switch (component) {
            case "month": date.Month = 13; break;
            case "day": date.Day = 32; break;
            case "end-month": date.EndYear = 2025; date.EndMonth = 13; break;
            case "end-day": date.EndYear = 2025; date.EndMonth = 1; date.EndDay = 32; break;
        }
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Dates.Add(date);
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV218" && diagnostic.Field == "dates.Issued");
    }

    [Theory]
    [InlineData("ISBN")]
    [InlineData("ISSN")]
    public void EndNote_serial_identifiers_replace_invalid_XML_characters_only_in_permissive_output(string scheme) {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Identifiers.Add(new BibliographyIdentifier(scheme, "before\u0001after"));
        document.Items.Add(item);

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        BibliographyIdentifier reopened = Assert.Single(BibliographyDocument.Parse(permissive.Content, BibliographyFormat.EndNoteXml).Document.Items[0].Identifiers);

        Assert.Contains(permissive.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV210" && diagnostic.Field == "identifiers." + scheme);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV210" && diagnostic.Field == "identifiers." + scheme);
        Assert.Equal("before\uFFFDafter", reopened.Value);
    }

    [Theory]
    [InlineData("root", "xml")]
    [InlineData("records", "records")]
    [InlineData("record", "@record-attributes")]
    public void EndNote_interleaved_namespace_attributes_preserve_their_carrier_order(string owner, string field) {
        string rootAttributes = owner == "root" ? " first=\"1\" xmlns:p=\"urn:p\" p:second=\"2\"" : string.Empty;
        string recordsAttributes = owner == "records" ? " first=\"1\" xmlns:p=\"urn:p\" p:second=\"2\"" : string.Empty;
        string recordAttributes = owner == "record" ? " first=\"1\" xmlns:p=\"urn:p\" p:second=\"2\"" : string.Empty;
        string source = "<xml" + rootAttributes + "><records" + recordsAttributes + "><record" + recordAttributes + "><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        string originalCarrier = owner == "record"
            ? Assert.Single(document.Items[0].NativeFields, native => native.Name == field).Value
            : Assert.Single(document.NativeEntries, native => native.Kind == "attributes" && native.Name == field).Value;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDocument reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document;
        string reopenedCarrier = owner == "record"
            ? Assert.Single(reopened.Items[0].NativeFields, native => native.Name == field).Value
            : Assert.Single(reopened.NativeEntries, native => native.Kind == "attributes" && native.Name == field).Value;

        Assert.Equal(originalCarrier, reopenedCarrier);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Action != BibliographyConversionAction.PreservedExtension);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Preserve_mode_rejects_unsupported_EndNote_XML_encoding_declarations(bool requireNoLoss) {
        const string source = "<?xml version=\"1.0\" encoding=\"x-custom\"?><xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Preserve, RequireNoLoss = requireNoLoss }));

        Assert.Contains("x-custom", exception.Message, StringComparison.Ordinal);
    }
}
