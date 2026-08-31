namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave59RegressionTests {
    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Bib_literal_name_wrappers_diagnose_terminal_backslash_normalization(BibliographyFormat format) {
        BibliographyDocument document = CreateBibDocument(format);
        document.Items[0].Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Literal = "Example\\" }));

        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV134" && diagnostic.Field == "author");
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Bib_keyword_wrappers_diagnose_terminal_backslash_normalization(BibliographyFormat format) {
        BibliographyDocument document = CreateBibDocument(format);
        document.Items[0].Keywords.Add("alpha,beta\\");

        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV134" && diagnostic.Field == "keywords");
    }

    [Theory]
    [InlineData("element")]
    [InlineData("records-element")]
    public void EndNote_extension_elements_diagnose_literal_attribute_whitespace(string kind) {
        BibliographyDocument document = CreateEndNoteDocument();
        document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, kind, "<extra label=\"A\tB\"/>", "extra"));

        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV251" && diagnostic.Field == "extra");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Edited_EndNote_native_XML_parsing_observes_cancellation(bool directCodec) {
        string originalValue = new string('x', 64 * 1024 * 1024);
        BibliographyDocument document = CreateEndNoteDocument();
        var field = new BibliographyNativeField(BibliographyFormat.EndNoteXml, "custom", originalValue, "<custom>" + originalValue + "</custom>");
        field.Value = "edited";
        document.Items[0].NativeFields.Add(field);
        using var cancellation = new CancellationTokenSource();
        var cancellationThread = new Thread(() => { Thread.Sleep(1); cancellation.Cancel(); });
        cancellationThread.Start();

        try {
            Assert.Throws<OperationCanceledException>(() => {
                if (directCodec) EndNoteXmlCodec.Write(document, new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, new BibliographyConversionReport(), cancellation.Token);
                else document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, cancellation.Token);
            });
        } finally {
            cancellationThread.Join();
        }
    }

    [Theory]
    [InlineData("xml", "other", "BIBCONV131")]
    [InlineData("records", "other", "BIBCONV131")]
    [InlineData("@record-attributes", "other", "BIBCONV132")]
    [InlineData("xml", "xml", "BIBCONV131")]
    [InlineData("records", "records", "BIBCONV131")]
    [InlineData("@record-attributes", "record", "BIBCONV132")]
    public void EndNote_attribute_carriers_require_the_owner_name_and_attributes_only(string owner, string carrierName, string diagnosticCode) {
        BibliographyDocument document = CreateEndNoteDocument();
        string carrier = "<" + carrierName + " custom=\"kept\"><lost/></" + carrierName + ">";
        if (owner == "@record-attributes") document.Items[0].NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, owner, carrier));
        else document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, "attributes", carrier, owner));

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.DoesNotContain("custom=\"kept\"", permissive.Content, StringComparison.Ordinal);
        Assert.Contains(permissive.Report.Diagnostics, diagnostic => diagnostic.Code == diagnosticCode);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == diagnosticCode);
    }

    [Theory]
    [InlineData("xml")]
    [InlineData("records")]
    [InlineData("@record-attributes")]
    public void Empty_matching_EndNote_attribute_carriers_remain_exact(string owner) {
        BibliographyDocument document = CreateEndNoteDocument();
        string carrierName = owner == "@record-attributes" ? "record" : owner;
        string carrier = "<" + carrierName + " custom=\"kept\"/>";
        if (owner == "@record-attributes") document.Items[0].NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, owner, carrier));
        else document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, "attributes", carrier, owner));

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains("custom=\"kept\"", written.Content, StringComparison.Ordinal);
    }

    private static BibliographyDocument CreateBibDocument(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Title" });
        return document;
    }

    private static BibliographyDocument CreateEndNoteDocument() {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Title" });
        return document;
    }
}
