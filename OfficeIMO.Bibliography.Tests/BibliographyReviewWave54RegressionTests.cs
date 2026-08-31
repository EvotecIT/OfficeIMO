namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave54RegressionTests {
    [Theory]
    [InlineData("element")]
    [InlineData("records-element")]
    public void EndNote_document_extensions_diagnose_carriage_return_normalization(string kind) {
        BibliographyDocument document = CreateEndNoteDocument();
        document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, kind, "<metadata>A\rB</metadata>", "metadata"));

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        BibliographyNativeEntry reopened = Assert.Single(BibliographyDocument.Parse(permissive.Content, BibliographyFormat.EndNoteXml).Document.NativeEntries,
            entry => entry.Kind == kind && entry.Name == "metadata");

        Assert.Contains("A\nB", reopened.Value, StringComparison.Ordinal);
        Assert.DoesNotContain('\r', reopened.Value);
        Assert.Contains(permissive.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV235" && diagnostic.Field == "metadata");
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV235" && diagnostic.Field == "metadata");
    }

    [Theory]
    [InlineData("element")]
    [InlineData("records-element")]
    public void EndNote_document_extension_parsing_and_writing_observe_cancellation(string kind) {
        BibliographyDocument document = CreateEndNoteDocument();
        document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, kind, "<metadata>" + new string('x', 64 * 1024 * 1024) + "</metadata>", "metadata"));
        using var cancellation = new CancellationTokenSource();
        var cancellationThread = new Thread(() => { Thread.Sleep(1); cancellation.Cancel(); });
        cancellationThread.Start();

        try {
            Assert.Throws<OperationCanceledException>(() =>
                EndNoteXmlCodec.Write(document, new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, new BibliographyConversionReport(), cancellation.Token));
        } finally {
            cancellationThread.Join();
        }
    }

    private static BibliographyDocument CreateEndNoteDocument() {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Title" });
        return document;
    }
}
