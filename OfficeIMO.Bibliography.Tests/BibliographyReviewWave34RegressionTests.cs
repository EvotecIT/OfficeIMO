using System.Collections.Generic;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave34RegressionTests {
    [Fact]
    public async Task Async_stream_save_checks_cancellation_before_truncating_a_seekable_destination() {
        byte[] original = Encoding.UTF8.GetBytes("existing destination");
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancelOnCanSeekStream(original, cancellation);
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Replacement" });
        stream.ArmCancellation();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            document.SaveAsync(stream, new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, cancellation.Token));

        Assert.Equal(original, stream.ToArray());
    }


    [Fact]
    public void Baseline_fingerprinting_observes_cancellation_within_a_large_value() {
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = new string('x', 64 * 1024 * 1024) };
        using var cancellation = new CancellationTokenSource();
        var cancellationThread = new Thread(() => { Thread.Sleep(1); cancellation.Cancel(); });
        cancellationThread.Start();

        try {
            Assert.Throws<OperationCanceledException>(() =>
                new BibliographyDocument(
                    BibliographyFormat.CslJson,
                    new List<BibliographyItem> { item },
                    new List<BibliographyNativeEntry>(),
                    "[]",
                    null,
                    Array.Empty<BibliographyDiagnostic>(),
                    cancellationToken: cancellation.Token));
        } finally {
            cancellationThread.Join();
        }
    }

    [Theory]
    [InlineData("\r")]
    [InlineData("\r\n")]
    public void Bib_diagnostic_locations_recognize_carriage_return_line_endings(string lineEnding) {
        string source = "@book{x}" + lineEnding + "  ignored" + lineEnding + "@book{y}";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibTex);

        BibliographyDiagnostic diagnostic = Assert.Single(read.Diagnostics, candidate => candidate.Code == "BIBBIB001");
        Assert.Equal(2, diagnostic.Line);
        Assert.Equal(3, diagnostic.Column);
        Assert.Equal(source.IndexOf("ignored", StringComparison.Ordinal), diagnostic.Offset);
    }

    private sealed class CancelOnCanSeekStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        private bool _armed;

        internal CancelOnCanSeekStream(byte[] source, CancellationTokenSource cancellation) : base(source.Length + 1024) {
            _cancellation = cancellation;
            Write(source, 0, source.Length);
        }

        public override bool CanSeek {
            get {
                if (_armed) {
                    _armed = false;
                    _cancellation.Cancel();
                }
                return base.CanSeek;
            }
        }

        internal void ArmCancellation() => _armed = true;
    }

}
