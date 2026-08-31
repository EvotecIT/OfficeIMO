using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave46RegressionTests {
    [Fact]
    public void CSL_contributor_projection_observes_cancellation_after_JSON_materialization() {
        string source = "[" + string.Join(",", Enumerable.Repeat("\"Name\"", 500_000)) + "]";
        using JsonDocument json = JsonDocument.Parse(source);
        var item = new BibliographyItem();
        var items = new List<BibliographyItem> { item };
        var limits = new BibliographyLimitGuard(new BibliographyReadOptions());
        using var cancellation = new CancellationTokenSource();
        var cancellationThread = new Thread(() => { Thread.Sleep(1); cancellation.Cancel(); });
        cancellationThread.Start();

        try {
            Assert.Throws<OperationCanceledException>(() => {
                CslJsonCodec.ParseNames(item, json.RootElement, BibliographyContributorRole.Author, items, limits, cancellation.Token);
            });
        } finally {
            cancellationThread.Join();
        }
    }

    [Fact]
    public void EndNote_byte_loading_scans_the_complete_XML_declaration_for_legacy_encoding() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
        string source = "<?xml version=\"1.0\"" + new string(' ', 8192) + "encoding=\"windows-1252\"?><xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Café €</title></titles></record></records></xml>";
        using var stream = new MemoryStream(windows1252.GetBytes(source));

        BibliographyReadResult read = BibliographyDocument.Load(stream, BibliographyFormat.EndNoteXml);

        Assert.False(read.HasErrors);
        Assert.Equal("Café €", Assert.Single(read.Document.Items).Title);
    }

    [Fact]
    public void EndNote_text_preservation_scans_the_complete_XML_declaration_for_legacy_encoding() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
        string source = "<?xml version=\"1.0\"" + new string(' ', 8192) + "encoding=\"windows-1252\"?><xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Café €</title></titles></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;

        BibliographyWriteResult written = document.Write();

        Assert.True(written.UsedOriginalSource);
        Assert.Equal(windows1252.GetBytes(source), written.Bytes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BOMless_UTF16_detection_scans_beyond_the_old_prefix(bool bigEndian) {
        string source = new string(' ', 8192) + "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Detected</title></titles></record></records></xml>";
        var encoding = new UnicodeEncoding(bigEndian, false, true);
        using var stream = new MemoryStream(encoding.GetBytes(source));

        BibliographyReadResult read = BibliographyDocument.Load(stream, BibliographyFormat.EndNoteXml);

        Assert.False(read.HasErrors);
        Assert.Equal("Detected", Assert.Single(read.Document.Items).Title);
    }

    [Fact]
    public void Long_XML_declaration_scans_observe_cancellation() {
        byte[] bytes = Encoding.ASCII.GetBytes("<?xml" + new string(' ', 16 * 1024 * 1024));
        using var cancellation = new CancellationTokenSource();
        var cancellationThread = new Thread(() => { Thread.Sleep(1); cancellation.Cancel(); });
        cancellationThread.Start();

        try {
            Assert.Throws<OperationCanceledException>(() => BibliographyEncoding.Detect(bytes, cancellation.Token));
        } finally {
            cancellationThread.Join();
        }
    }
}
