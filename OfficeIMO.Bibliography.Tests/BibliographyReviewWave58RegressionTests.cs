namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave58RegressionTests {
    [Theory]
    [InlineData("element")]
    [InlineData("records-element")]
    public void EndNote_native_entries_reject_trailing_XML_content(string kind) {
        var document = CreateDocument();
        document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, kind, "<extra>kept</extra><lost>discarded</lost>", "extra"));

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.DoesNotContain("<extra>", permissive.Content, StringComparison.Ordinal);
        Assert.Contains(permissive.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV117" && diagnostic.Field == "extra");
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV117" && diagnostic.Field == "extra");
    }

    [Theory]
    [InlineData("xml", "BIBCONV131")]
    [InlineData("records", "BIBCONV131")]
    [InlineData("@record-attributes", "BIBCONV132")]
    public void EndNote_attribute_carriers_reject_trailing_XML_content(string owner, string diagnosticCode) {
        var document = CreateDocument();
        string elementName = owner == "@record-attributes" ? "record" : owner;
        string carrier = "<" + elementName + " custom=\"kept\"/><lost/>";
        if (owner == "@record-attributes") document.Items[0].NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, owner, carrier));
        else document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, "attributes", carrier, owner));

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.DoesNotContain("custom=\"kept\"", permissive.Content, StringComparison.Ordinal);
        Assert.Contains(permissive.Report.Diagnostics, diagnostic => diagnostic.Code == diagnosticCode);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == diagnosticCode);
    }

    [Fact]
    public void Edited_EndNote_native_fields_use_the_public_name_instead_of_a_mismatched_raw_element() {
        var document = CreateDocument();
        var field = new BibliographyNativeField(BibliographyFormat.EndNoteXml, "expected", "old", "<other custom=\"retained\">old</other>");
        field.Value = "new";
        document.Items[0].NativeFields.Add(field);

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        BibliographyNativeField reopened = Assert.Single(BibliographyDocument.Parse(permissive.Content, BibliographyFormat.EndNoteXml).Document.Items[0].NativeFields,
            candidate => candidate.Name == "expected");

        Assert.Equal("new", reopened.Value);
        Assert.Contains("<expected>new</expected>", permissive.Content, StringComparison.Ordinal);
        Assert.DoesNotContain("<other", permissive.Content, StringComparison.Ordinal);
        Assert.Contains(permissive.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV247" && diagnostic.Field == "native.expected");
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV247" && diagnostic.Field == "native.expected");
    }

    [Fact]
    public void Edited_EndNote_native_fields_keep_matching_raw_element_metadata() {
        var document = CreateDocument();
        var field = new BibliographyNativeField(BibliographyFormat.EndNoteXml, "expected", "old", "<m:expected xmlns:m=\"urn:metadata\" custom=\"retained\">old</m:expected>");
        field.Value = "new";
        document.Items[0].NativeFields.Add(field);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains("<m:expected", written.Content, StringComparison.Ordinal);
        Assert.Contains("custom=\"retained\"", written.Content, StringComparison.Ordinal);
        Assert.Contains(">new</m:expected>", written.Content, StringComparison.Ordinal);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV247");
    }

    private static BibliographyDocument CreateDocument() {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Title" });
        return document;
    }
}
