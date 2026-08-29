namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave19RegressionTests {
    [Fact]
    public void Structured_Bib_names_with_component_commas_block_strict_output() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Names" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Doe, Smith", Given = "Jane" }));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV226" && diagnostic.Field == "contributors");
    }

    [Theory]
    [InlineData("title", "123")]
    [InlineData("publisher", "true")]
    [InlineData("DOI", "123")]
    [InlineData("keyword", "false")]
    [InlineData("note", "123")]
    [InlineData("type", "123")]
    public void Non_string_CSL_scalars_remain_native_JSON(string property, string rawValue) {
        string typedType = property == "type" ? string.Empty : ",\"type\":\"book\"";
        string source = "[{\"id\":\"x\"" + typedType + ",\"" + property + "\":" + rawValue + "}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        BibliographyNativeField field = Assert.Single(reopened.NativeFields, field => field.Format == BibliographyFormat.CslJson && field.Name == property);
        Assert.Equal(rawValue, field.RawValue);
    }

    [Fact]
    public void Non_string_CSL_name_components_remain_native_JSON() {
        const string source = "[{\"id\":\"x\",\"type\":\"book\",\"author\":[{\"family\":123}]}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        BibliographyNativeField field = Assert.Single(reopened.NativeFields, field => field.Name == "author");
        using System.Text.Json.JsonDocument raw = System.Text.Json.JsonDocument.Parse(field.RawValue!);
        Assert.Equal(System.Text.Json.JsonValueKind.Number, raw.RootElement[0].GetProperty("family").ValueKind);
        Assert.Equal(123, raw.RootElement[0].GetProperty("family").GetInt32());
    }

    [Fact]
    public void Non_string_CSL_date_literals_remain_native_JSON() {
        const string source = "[{\"id\":\"x\",\"type\":\"book\",\"issued\":{\"literal\":123}}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate date = Assert.Single(Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items).Dates);

        BibliographyNativeField field = Assert.Single(date.NativeFields, field => field.Name == "literal");
        Assert.Equal("123", field.RawValue);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void Empty_Bib_keywords_remain_distinct_from_an_absent_list(string keyword) {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Keywords" };
        item.Keywords.Add(keyword);
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibLatex).Document.Items);

        Assert.Equal(keyword, Assert.Single(reopened.Keywords));
    }

    [Fact]
    public void Invalid_Bib_text_observes_cancellation_during_recovery_scan() {
        string source = new string('x', 32 * 1024 * 1024);
        using var cancellation = new CancellationTokenSource();
        cancellation.CancelAfter(TimeSpan.FromMilliseconds(10));

        Assert.Throws<OperationCanceledException>(() =>
            BibliographyDocument.Parse(source, BibliographyFormat.BibLatex, cancellationToken: cancellation.Token));
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.CslJson)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Generated_citation_keys_do_not_collide_with_existing_keys(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        document.Items.Add(new BibliographyItem { Type = BibliographyItemType.Book, Title = "Generated" });
        document.Items.Add(new BibliographyItem { Key = "item-1", Type = BibliographyItemType.Book, Title = "Existing" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        string?[] keys = BibliographyDocument.Parse(written.Content, format).Document.Items.Select(static item => item.Key).ToArray();

        Assert.Equal(2, keys.Distinct(StringComparer.OrdinalIgnoreCase).Count());
    }

    [Fact]
    public void EndNote_records_extensions_cannot_promote_into_typed_records() {
        const string source = "<xml><records><metadata>retained</metadata><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        BibliographyNativeEntry entry = Assert.Single(document.NativeEntries, entry => entry.Kind == "records-element");
        entry.Value = "<record><rec-number>2</rec-number><ref-type name=\"Book\">6</ref-type></record>";

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV117" && diagnostic.Field == "metadata");
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    public void Malformed_tagged_lines_report_their_start_offset(BibliographyFormat format) {
        BibliographyReadResult read = BibliographyDocument.Parse("\nbad line", format);

        BibliographyDiagnostic diagnostic = Assert.Single(read.Diagnostics, diagnostic => diagnostic.Code == "BIBTAG001");
        Assert.Equal(1, diagnostic.Offset);
        Assert.Equal(2, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }
}
