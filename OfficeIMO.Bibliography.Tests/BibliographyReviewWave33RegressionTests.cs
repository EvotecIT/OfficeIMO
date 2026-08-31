namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave33RegressionTests {
    [Theory]
    [InlineData("é")]
    [InlineData("😀")]
    public void CSL_limit_diagnostics_use_zero_based_UTF16_offsets(string prefix) {
        string source = "[{\"id\":\"" + prefix + "\",\"x\":\"long\"}]";
        var options = new BibliographyReadOptions { MaximumValueLength = 3 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, options);

        BibliographyDiagnostic diagnostic = Assert.Single(read.Diagnostics, candidate => candidate.Code == "BIBLIM001");
        Assert.Equal(source.IndexOf("\"long\"", StringComparison.Ordinal), diagnostic.Offset);
    }

    [Fact]
    public void CSL_limit_diagnostics_include_a_leading_BOM_in_UTF16_offsets() {
        string source = "\uFEFF[{\"id\":\"é\",\"x\":\"long\"}]";
        var options = new BibliographyReadOptions { MaximumValueLength = 3 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, options);

        BibliographyDiagnostic diagnostic = Assert.Single(read.Diagnostics, candidate => candidate.Code == "BIBLIM001");
        Assert.Equal(source.IndexOf("\"long\"", StringComparison.Ordinal), diagnostic.Offset);
    }

    [Fact]
    public void CSL_large_token_parsing_observes_cancellation() {
        string source = "[{\"id\":\"x\",\"title\":\"" + new string('x', 16 * 1024 * 1024) + "\"}]";
        var options = new BibliographyReadOptions { MaximumValueLength = 20 * 1024 * 1024 };
        BibliographyCancellationTest.AssertObserved(token =>
            BibliographyDocument.Parse(source, BibliographyFormat.CslJson, options, token));
    }

    [Fact]
    public void CSL_decoded_length_accounting_handles_an_escape_split_across_input_segments() {
        const string opening = "[{\"id\":\"x\",\"title\":\"";
        int fillerLength = 4095 - Encoding.UTF8.GetByteCount(opening);
        string source = opening + new string('a', fillerLength) + "\\u00e9\"}]";
        var options = new BibliographyReadOptions { MaximumValueLength = fillerLength + 1 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, options);

        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
        Assert.Equal(new string('a', fillerLength) + "é", read.Document.Items[0].Title);
    }

    [Fact]
    public void Synchronous_stream_save_observes_cancellation_between_bounded_writes() {
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancelAfterFirstWriteStream(cancellation);
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = new string('x', 256 * 1024) });

        Assert.Throws<OperationCanceledException>(() =>
            document.Save(stream, new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, cancellation.Token));

        Assert.Equal(1, stream.WriteCount);
        Assert.InRange(stream.MaximumWriteSize, 1, 81920);
    }

    [Theory]
    [InlineData("doi", "10.1000/example")]
    [InlineData("IsBn", "978-1-4028-9462-6")]
    [InlineData("AccessIon", "example")]
    public void RIS_identifier_scheme_case_normalization_is_diagnosed(string scheme, string value) {
        BibliographyDocument document = CreateDocument(BibliographyFormat.Ris, scheme, value);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV228" && diagnostic.Field == "identifiers." + scheme);
    }

    [Theory]
    [InlineData("pmid", "1")]
    [InlineData("issn", "1234-5678")]
    public void NBIB_standard_identifier_scheme_case_normalization_is_diagnosed(string scheme, string value) {
        BibliographyDocument document = CreateDocument(BibliographyFormat.Nbib, scheme, value);
        if (!string.Equals(scheme, "pmid", StringComparison.OrdinalIgnoreCase)) document.Items[0].Identifiers.Insert(0, new BibliographyIdentifier("PMID", "1"));

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV232" && diagnostic.Field == "identifiers." + scheme);
    }

    [Theory]
    [InlineData("doi")]
    [InlineData("AccessIon")]
    public void EndNote_identifier_scheme_case_normalization_is_diagnosed(string scheme) {
        BibliographyDocument document = CreateDocument(BibliographyFormat.EndNoteXml, scheme, "example");

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV204" && diagnostic.Field == "identifiers." + scheme);
    }

    [Fact]
    public void EndNote_serial_identifier_scheme_case_is_preserved_explicitly() {
        BibliographyDocument document = CreateDocument(BibliographyFormat.EndNoteXml, "IsBn", "978-1-4028-9462-6");

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyIdentifier reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items[0].Identifiers);

        Assert.Equal("IsBn", reopened.Scheme);
    }

    private static BibliographyDocument CreateDocument(BibliographyFormat format, string scheme, string value) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Identifiers.Add(new BibliographyIdentifier(scheme, value));
        document.Items.Add(item);
        return document;
    }

    private sealed class CancelAfterFirstWriteStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        internal CancelAfterFirstWriteStream(CancellationTokenSource cancellation) => _cancellation = cancellation;
        internal int WriteCount { get; private set; }
        internal int MaximumWriteSize { get; private set; }

        public override void Write(byte[] buffer, int offset, int count) {
            WriteCount++;
            MaximumWriteSize = Math.Max(MaximumWriteSize, count);
            base.Write(buffer, offset, count);
            if (WriteCount == 1) _cancellation.Cancel();
        }
    }
}
