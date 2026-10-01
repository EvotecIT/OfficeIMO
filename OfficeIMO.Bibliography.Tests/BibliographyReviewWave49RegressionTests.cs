namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave49RegressionTests {
    [Theory]
    [InlineData("TY  - BOOK\nDO  - earlier\nAN  - DOI:10.1/\n      suffix\nER  - \n", "DOI:10.1/ suffix", "DOI", "earlier")]
    [InlineData("TY  - BOOK\nAN  - local:first\nAN  - DOI:10.1/\n      suffix\nER  - \n", "local:first", "local", "first")]
    public void RIS_accession_continuations_update_the_exact_qualified_identifier(string source, string expectedKey, string earlierScheme, string earlierValue) {
        BibliographyItem item = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document.Items);

        Assert.Equal(expectedKey, item.Key);
        Assert.Equal(new[] { earlierScheme + "=" + earlierValue, "DOI=10.1/ suffix" }, item.Identifiers.Select(identifier => identifier.Scheme + "=" + identifier.Value));
    }


    [Theory]
    [InlineData("validate")]
    [InlineData("sanitize")]
    public void EndNote_native_value_XML_scans_observe_cancellation_during_large_values(string operation) {
        string value = new string('x', 64 * 1024 * 1024);
        using var cancellation = new CancellationTokenSource();
        var cancellationThread = new Thread(() => { Thread.Sleep(1); cancellation.Cancel(); });
        cancellationThread.Start();

        try {
            Assert.Throws<OperationCanceledException>(() => {
                if (operation == "validate") EndNoteXmlCodec.HasInvalidXmlCharacters(value, cancellation.Token);
                else EndNoteXmlCodec.SanitizeXml(value, cancellation.Token);
            });
        } finally {
            cancellationThread.Join();
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Caller_selected_UTF8_rejects_malformed_bytes_without_mutating_the_encoding(bool asynchronous) {
        byte[] bytes = { (byte)'T', (byte)'Y', (byte)' ', (byte)' ', (byte)'-', (byte)' ', (byte)'B', (byte)'O', (byte)'O', (byte)'K', (byte)'\n', (byte)'T', (byte)'I', (byte)' ', (byte)' ', (byte)'-', (byte)' ', 0xC3, 0x28 };
        Encoding encoding = Encoding.UTF8;
        DecoderFallback originalFallback = encoding.DecoderFallback;
        using var stream = new MemoryStream(bytes);

        if (asynchronous) {
            await Assert.ThrowsAsync<DecoderFallbackException>(() =>
                BibliographyDocument.LoadAsync(stream, BibliographyFormat.Ris, encoding: encoding));
        } else {
            Assert.Throws<DecoderFallbackException>(() =>
                BibliographyDocument.Load(stream, BibliographyFormat.Ris, encoding: encoding));
        }

        Assert.Same(originalFallback, encoding.DecoderFallback);
    }

    [Fact]
    public void Caller_selected_multibyte_encodings_reject_incomplete_terminal_sequences() {
        Encoding[] encodings = {
            new UnicodeEncoding(false, false),
            new UnicodeEncoding(true, false),
            new UTF32Encoding(false, false),
            new UTF32Encoding(true, false)
        };

        foreach (Encoding encoding in encodings)
            Assert.Throws<DecoderFallbackException>(() => BibliographyEncoding.DecodeBounded(new byte[] { 0x41 }, encoding, 16, CancellationToken.None));
    }
}
