using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfForwardOnlyObjectSerializationTests {
    [Fact]
    public void ForwardOnlyObjectSerialization_WritesReadableDeterministicPdfWithoutObjectReplay() {
        PdfOptions options = CreateForwardOnlyOptions();
        byte[] first;
        PdfSerializationReport report;
        {
            PdfDocument document = PdfDocument.Create(options)
            .H1("Forward-only objects")
            .Paragraph(paragraph => paragraph.Text("Objects are emitted once; layout remains truthfully replayable."));
            using var output = new MemoryStream();
            PdfSaveResult save = document.Save(output);
            first = output.ToArray();
            report = Assert.IsType<PdfSerializationReport>(save.Serialization);
        }

        byte[] second = PdfDocument.Create(CreateForwardOnlyOptions())
            .H1("Forward-only objects")
            .Paragraph(paragraph => paragraph.Text("Objects are emitted once; layout remains truthfully replayable."))
            .ToBytes();

        Assert.Equal(first, second);
        Assert.Contains("Forward-only objects", PdfReadDocument.Open(first).ExtractText(), StringComparison.Ordinal);
        Assert.True(report.IsForwardOnlyObjectSerialization);
        Assert.False(report.IsForwardOnlyLayout);
        Assert.Equal(0L, report.PeakRetainedObjectBytes);
        Assert.False(report.ObjectBufferSpilled);
        Assert.True(report.LargestSerializedObjectBytes > 0L);
    }

    [Fact]
    public void ForwardOnlyObjectSerialization_SupportsNonSeekableDestination() {
        using var destination = new NonSeekableWriteStream();
        PdfDocument document = PdfDocument.Create(CreateForwardOnlyOptions())
            .Paragraph(paragraph => paragraph.Text("Non-seekable destination"));

        PdfSaveResult save = document.Save(destination);
        byte[] bytes = destination.ToArray();

        Assert.True(save.Succeeded);
        Assert.True(save.Serialization?.IsForwardOnlyObjectSerialization);
        Assert.Contains("Non-seekable destination", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void ForwardOnlyObjectSerialization_RequiresExplicitModernHeader() {
        var options = new PdfOptions {
            ObjectSerializationMode = PdfObjectSerializationMode.ForwardOnly
        };
        PdfDocument document = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Invalid profile"));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => document.ToBytes());

        Assert.Contains("PDF 1.7", exception.Message, StringComparison.Ordinal);
    }

    private static PdfOptions CreateForwardOnlyOptions() =>
        new() {
            FileVersion = PdfFileVersion.Pdf17,
            ObjectSerializationMode = PdfObjectSerializationMode.ForwardOnly,
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers
        };

    private sealed class NonSeekableWriteStream : Stream {
        private readonly MemoryStream _inner = new();
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => _inner.Length;
        public override long Position {
            get => _inner.Position;
            set => throw new NotSupportedException();
        }

        internal byte[] ToArray() => _inner.ToArray();
        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
