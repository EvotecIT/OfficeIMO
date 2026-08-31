namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave56RegressionTests {
    [Theory]
    [InlineData("root", "xml")]
    [InlineData("records", "records")]
    [InlineData("record", "@record-attributes")]
    public void EndNote_attribute_carriers_diagnose_literal_whitespace_normalization(string owner, string field) {
        const string value = "A\tB\nC\rD";
        string carrier = "<" + (owner == "record" ? "record" : field) + " custom=\"" + value + "\" />";
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Title" };
        document.Items.Add(item);
        if (owner == "record") item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, field, carrier));
        else document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, "attributes", carrier, field));

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        BibliographyDocument reopened = BibliographyDocument.Parse(permissive.Content, BibliographyFormat.EndNoteXml).Document;
        string reopenedCarrier = owner == "record"
            ? Assert.Single(reopened.Items[0].NativeFields, candidate => candidate.Name == field).Value
            : Assert.Single(reopened.NativeEntries, candidate => candidate.Kind == "attributes" && candidate.Name == field).Value;

        Assert.Contains("custom=\"A B C D\"", reopenedCarrier, StringComparison.Ordinal);
        Assert.Contains(permissive.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV251" && diagnostic.Field == field);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV251" && diagnostic.Field == field);
    }

    [Fact]
    public void EndNote_attribute_character_references_remain_exact() {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, "attributes", "<xml custom=\"A&#x9;B&#xA;C&#xD;D\" />", "xml"));
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Title" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyNativeEntry reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.NativeEntries,
            candidate => candidate.Kind == "attributes" && candidate.Name == "xml");

        Assert.Contains("custom=\"A&#x9;B&#xA;C&#xD;D\"", reopened.Value, StringComparison.Ordinal);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV251");
    }
}
