namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave27RegressionTests {
    [Theory]
    [InlineData("contributors")]
    [InlineData("titles")]
    [InlineData("periodical")]
    [InlineData("dates")]
    [InlineData("urls")]
    [InlineData("keywords")]
    public void Empty_EndNote_record_containers_survive_unrelated_strict_edits(string container) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><" + container + "/></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Publisher = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Contains(reopened.NativeFields, field => field.Format == BibliographyFormat.EndNoteXml && string.Equals(field.Name, container, StringComparison.OrdinalIgnoreCase) && field.Value.Length == 0);
    }

    [Theory]
    [InlineData("isbn", "978-1-4028-9462-6")]
    [InlineData("issn", "2049-3630")]
    public void EndNote_identifier_scheme_casing_survives_strict_edits(string scheme, string value) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles><isbn type=\"" + scheme + "\">" + value + "</isbn></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyIdentifier reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items[0].Identifiers);

        Assert.Equal(scheme, reopened.Scheme);
        Assert.Contains("type=\"" + scheme + "\"", written.Content, StringComparison.Ordinal);
    }

    [Fact]
    public void Classic_BibTeX_month_only_dates_are_valid_partial_dates() {
        BibliographyDocument document = BibliographyDocument.Parse("@book{x, month={January}, title={Before}}", BibliographyFormat.BibTex).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibTex).Document.Items[0].Dates);

        Assert.Equal(1, reopened.Month);
        Assert.Null(reopened.Year);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV218");
    }

    [Fact]
    public void Conflicting_EndNote_type_name_and_number_are_diagnosed_as_recovered_loss() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">17</ref-type><titles><title>Before</title></titles></record></records></xml>";
        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml);
        read.Document.Items[0].Title = "After";

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBEND004" && diagnostic.Field == "ref-type");
        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            read.Document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV222" && diagnostic.Field == "ref-type");
    }

    [Fact]
    public void EndNote_offset_mapping_is_constant_memory_and_exact_for_many_lines() {
        string source = new string('\n', 2_000_000);
#if NET472
        var offsets = new EndNoteSourceOffsetMap(source, 0, CancellationToken.None);
#else
        long before = GC.GetAllocatedBytesForCurrentThread();
        var offsets = new EndNoteSourceOffsetMap(source, 0, CancellationToken.None);
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.True(allocated < 128 * 1024, $"EndNote offset mapping allocated {allocated:N0} bytes before lookup.");
#endif

        Assert.Equal(source.Length, offsets.GetOffset(new FixedLineInfo(source.Length + 1, 1)));
        Assert.Equal(1, offsets.GetOffset(new FixedLineInfo(2, 1)));
    }

    [Fact]
    public void EndNote_offset_mapping_observes_cancellation_during_large_scans() {
        string source = new string('\n', 2_000_000);
        using var cancellation = new CancellationTokenSource();
        var offsets = new EndNoteSourceOffsetMap(source, 0, cancellation.Token);
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => offsets.GetOffset(new FixedLineInfo(source.Length + 1, 1)));
    }

    private sealed class FixedLineInfo : System.Xml.IXmlLineInfo {
        internal FixedLineInfo(int lineNumber, int linePosition) {
            LineNumber = lineNumber;
            LinePosition = linePosition;
        }

        public bool HasLineInfo() => true;
        public int LineNumber { get; }
        public int LinePosition { get; }
    }
}
