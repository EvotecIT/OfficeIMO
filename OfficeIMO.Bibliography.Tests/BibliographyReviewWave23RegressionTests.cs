using System.Globalization;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave23RegressionTests {
    [Fact]
    public void NBIB_author_normalization_is_linear_and_cancellation_aware() {
        var item = new BibliographyItem();
        for (int index = 0; index < 2_000; index++) {
            var full = new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Full" + index.ToString(CultureInfo.InvariantCulture), Given = "Person" });
            var compact = new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Other" + index.ToString(CultureInfo.InvariantCulture), Given = "P" });
            item.Contributors.Add(full);
            item.Contributors.Add(compact);
            item.TaggedContributorTags[full] = "FAU";
            item.TaggedContributorTags[compact] = "AU";
        }

#if NET472
        TaggedCodec.NormalizeNbibAuthors(new[] { item }, CancellationToken.None);
#else
        long before = GC.GetAllocatedBytesForCurrentThread();
        TaggedCodec.NormalizeNbibAuthors(new[] { item }, CancellationToken.None);
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.True(allocated < 20 * 1024 * 1024, $"NBIB author normalization allocated {allocated:N0} bytes for unmatched contributors.");
#endif
        Assert.Equal(4_000, item.Contributors.Count);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => TaggedCodec.NormalizeNbibAuthors(new[] { item }, cancellation.Token));
    }

    [Fact]
    public void NBIB_author_normalization_preserves_pairing_and_source_order() {
        const string source = "PMID- 1\nAU  - Smith J\nFAU - Smith, John\nFAU - Jones, Jane\nAU  - Jones J\nTI  - Authors\n";

        BibliographyItem item = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document.Items);

        Assert.Collection(item.Contributors,
            contributor => { Assert.Equal("Smith", contributor.Name.Family); Assert.Equal("John", contributor.Name.Given); },
            contributor => { Assert.Equal("Jones", contributor.Name.Family); Assert.Equal("Jane", contributor.Name.Given); });
    }

    [Fact]
    public void EndNote_root_extensions_cannot_promote_into_records_containers() {
        const string source = "<xml><extension>safe</extension><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        BibliographyNativeEntry extension = Assert.Single(document.NativeEntries, entry => entry.Kind == "element");
        extension.Value = "<records><record><rec-number>promoted</rec-number><ref-type name=\"Book\">6</ref-type></record></records>";

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV117" && diagnostic.Field == "extension");
    }

    [Fact]
    public void Foreign_namespace_EndNote_records_extensions_remain_safe() {
        const string source = "<xml xmlns:ext=\"urn:extension\"><ext:records><ext:record>native</ext:record></ext:records><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDocument reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document;

        Assert.Single(reopened.Items);
        Assert.Contains(reopened.NativeEntries, entry => entry.Kind == "element" && entry.Name == "records" && entry.Value.Contains("urn:extension", StringComparison.Ordinal));
    }

    [Fact]
    public void Empty_classic_BibTeX_dates_are_reported_before_they_disappear() {
        var document = new BibliographyDocument(BibliographyFormat.BibTex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued });
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV241" && diagnostic.Field == "dates.Issued");
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Canonical_Bib_edits_preserve_unknown_native_field_name_spelling(BibliographyFormat format) {
        BibliographyDocument document = BibliographyDocument.Parse("@Book{x,CustomField={value},title={Before}}", format).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items);

        Assert.Contains("CustomField = {value}", written.Content, StringComparison.Ordinal);
        Assert.Equal("CustomField", Assert.Single(reopened.NativeFields).Name);
    }

    [Fact]
    public void Nonempty_RIS_terminators_are_diagnosed_as_recovered_source_loss() {
        BibliographyReadResult read = BibliographyDocument.Parse("TY  - BOOK\nID  - x\nER  - checksum\n", BibliographyFormat.Ris);
        BibliographyDiagnostic parserDiagnostic = Assert.Single(read.Diagnostics, diagnostic => diagnostic.Code == "BIBTAG004");

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            read.Document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Equal("ER", parserDiagnostic.Field);
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV222" && diagnostic.Field == "ER");
    }
}
