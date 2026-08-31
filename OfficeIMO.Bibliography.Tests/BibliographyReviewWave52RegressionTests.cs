namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave52RegressionTests {
    [Fact]
    public void EndNote_native_types_with_invalid_XML_controls_are_sanitized_and_diagnosed() {
        AssertNativeTypeSanitized("Custom\u0001Type");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EndNote_native_types_with_unpaired_surrogates_are_sanitized_and_diagnosed(bool lowSurrogate) {
        AssertNativeTypeSanitized("Custom" + new string(lowSurrogate ? '\uDC00' : '\uD800', 1) + "Type");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EndNote_typed_text_with_unpaired_surrogates_is_sanitized_and_diagnosed(bool lowSurrogate) {
        string title = "Before" + new string(lowSurrogate ? '\uDC00' : '\uD800', 1) + "After";
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = title });

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(permissive.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal("Before\uFFFDAfter", reopened.Title);
        Assert.Contains(permissive.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV210" && diagnostic.Field == "title");
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV210" && diagnostic.Field == "title");
    }

    private static void AssertNativeTypeSanitized(string nativeType) {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Unknown, NativeType = nativeType });

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(permissive.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal("Custom\uFFFDType", reopened.NativeType);
        Assert.Contains(permissive.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV210" && diagnostic.Field == "type");
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV210" && diagnostic.Field == "type");
    }

    [Fact]
    public void EndNote_direct_records_roots_report_omitted_outer_root_extensions() {
        const string source = "<records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type></record></records>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, "element", "<metadata>retained</metadata>", "metadata"));

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.DoesNotContain("<metadata>", permissive.Content, StringComparison.Ordinal);
        Assert.Contains(permissive.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV116" && diagnostic.Field == "metadata");
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV116" && diagnostic.Field == "metadata");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EndNote_inherited_element_prefixes_survive_strict_canonical_edits(bool recordsRoot) {
        string source = recordsRoot
            ? "<e:records xmlns:e=\"urn:endnote\" xmlns:m=\"urn:metadata\" m:scope=\"all\"><e:record m:id=\"1\"><e:rec-number>1</e:rec-number><e:ref-type name=\"Book\">6</e:ref-type><e:titles><e:title>Before</e:title></e:titles></e:record></e:records>"
            : "<e:xml xmlns:e=\"urn:endnote\" xmlns:m=\"urn:metadata\"><e:records m:scope=\"all\"><e:record m:id=\"1\"><e:rec-number>1</e:rec-number><e:ref-type name=\"Book\">6</e:ref-type><e:titles><e:title>Before</e:title></e:titles></e:record></e:records></e:xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDocument reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document;

        Assert.Contains("<e:records", written.Content, StringComparison.Ordinal);
        Assert.Contains("<e:record", written.Content, StringComparison.Ordinal);
        Assert.Contains("m:scope=\"all\"", written.Content, StringComparison.Ordinal);
        Assert.Contains("m:id=\"1\"", written.Content, StringComparison.Ordinal);
        Assert.Equal("After", Assert.Single(reopened.Items).Title);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Action != BibliographyConversionAction.PreservedExtension);
    }
}
