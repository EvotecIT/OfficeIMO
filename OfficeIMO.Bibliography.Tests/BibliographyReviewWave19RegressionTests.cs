using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave19RegressionTests {

    [Theory]
    [InlineData("title", "123")]
    [InlineData("publisher", "true")]
    [InlineData("DOI", "123")]
    [InlineData("keyword", "false")]
    [InlineData("note", "123")]
    [InlineData("type", "123")]
    [InlineData("title", "{\"value\":\"Example\"}")]
    [InlineData("publisher", "{\"value\":\"Example\"}")]
    [InlineData("DOI", "{\"value\":\"Example\"}")]
    [InlineData("keyword", "{\"value\":\"Example\"}")]
    public void Non_string_CSL_scalars_remain_native_JSON(string property, string rawValue) {
        string typedType = property == "type" ? string.Empty : ",\"type\":\"book\"";
        string source = "[{\"id\":\"x\"" + typedType + ",\"" + property + "\":" + rawValue + "}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        BibliographyNativeField field = Assert.Single(reopened.NativeFields, field => field.Format == BibliographyFormat.CslJson && field.Name == property);
        Assert.Equal(rawValue, field.RawValue);
        if (rawValue.StartsWith("{", StringComparison.Ordinal)) {
            using JsonDocument json = JsonDocument.Parse(written.Content);
            JsonElement value = json.RootElement[0].GetProperty(property);
            Assert.Equal(JsonValueKind.Object, value.ValueKind);
            Assert.Equal("Example", value.GetProperty("value").GetString());
        }
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
