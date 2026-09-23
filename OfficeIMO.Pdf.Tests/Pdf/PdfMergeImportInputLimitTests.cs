using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfMergeImportInputLimitTests {
    [Fact]
    public void MergeAndImportStreamsBoundEverySourceIndependently() {
        byte[] validPdf = CreatePdf("valid source");

        using (var oversized = CreateOversizedStream()) {
            PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
                PdfMerger.Merge((IEnumerable<Stream>)new Stream[] { new MemoryStream(validPdf), oversized }));
            Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
            Assert.Equal(3, oversized.Position);
            Assert.False(oversized.WasRead);
        }

        using (var oversized = CreateOversizedStream()) {
            PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
                PdfMerger.Merge(new PdfMergeOptions(), (IEnumerable<Stream>)new Stream[] { oversized }));
            Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
            Assert.False(oversized.WasRead);
        }

        Action<Stream, Stream>[] importRoutes = {
            (target, source) => PdfPageImporter.AppendPages(target, source),
            (target, source) => PdfPageImporter.PrependPages(target, source),
            (target, source) => PdfPageImporter.InsertPages(target, source, 1),
            (target, source) => PdfPageImporter.InsertPageRange(target, source, 1, 1, 1),
            (target, source) => PdfPageImporter.InsertPageRange(target, source, 1, PdfPageRange.From(1, 1)),
            (target, source) => PdfPageImporter.AppendPageRanges(target, source, PdfPageRange.From(1, 1)),
            (target, source) => PdfPageImporter.PrependPageRanges(target, source, PdfPageRange.From(1, 1)),
            (target, source) => PdfPageImporter.InsertPageRanges(target, source, 1, PdfPageRange.From(1, 1))
        };

        foreach (Action<Stream, Stream> route in importRoutes) {
            using var target = new MemoryStream(validPdf);
            using var source = CreateOversizedStream();
            PdfReadLimitException sourceError = Assert.Throws<PdfReadLimitException>(() => route(target, source));
            Assert.Equal(PdfReadLimitKind.InputBytes, sourceError.Kind);
            Assert.Equal(3, source.Position);
            Assert.False(source.WasRead);

            using var oversizedTarget = CreateOversizedStream();
            using var unreadSource = new MemoryStream(validPdf);
            PdfReadLimitException targetError = Assert.Throws<PdfReadLimitException>(() => route(oversizedTarget, unreadSource));
            Assert.Equal(PdfReadLimitKind.InputBytes, targetError.Kind);
            Assert.Equal(3, oversizedTarget.Position);
            Assert.False(oversizedTarget.WasRead);
            Assert.Equal(0, unreadSource.Position);
        }
    }

    [Fact]
    public void MergeAndImportPathsRejectOversizedSourcesWithoutChangingOutputs() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-merge-import-limit-" + Guid.NewGuid().ToString("N"));
        string validPath = Path.Combine(root, "valid.pdf");
        string oversizedPath = Path.Combine(root, "oversized.pdf");
        string existingOutput = Path.Combine(root, "existing.pdf");
        string absentOutputDirectory = Path.Combine(root, "outputs");
        byte[] sentinel = { 7, 1, 7, 1 };
        try {
            Directory.CreateDirectory(root);
            File.WriteAllBytes(validPath, CreatePdf("valid path source"));
            using (var file = new FileStream(oversizedPath, FileMode.CreateNew, FileAccess.Write)) {
                file.SetLength(PdfLoadOptions.Default.Limits.MaxInputBytes + 1);
            }
            File.WriteAllBytes(existingOutput, sentinel);

            Assert.Equal(PdfReadLimitKind.InputBytes,
                Assert.Throws<PdfReadLimitException>(() => PdfMerger.MergeFilesToBytes(validPath, oversizedPath)).Kind);
            Assert.Equal(PdfReadLimitKind.InputBytes,
                Assert.Throws<PdfReadLimitException>(() => PdfMerger.MergeFilesToBytes(new PdfMergeOptions(), new[] { oversizedPath })).Kind);
            Assert.Equal(PdfReadLimitKind.InputBytes,
                Assert.Throws<PdfReadLimitException>(() => PdfPageImporter.AppendPages(validPath, oversizedPath)).Kind);
            Assert.Equal(PdfReadLimitKind.InputBytes,
                Assert.Throws<PdfReadLimitException>(() => PdfPageImporter.AppendPages(oversizedPath, validPath)).Kind);

            Assert.Throws<PdfReadLimitException>(() => PdfMerger.MergeFiles(new[] { validPath, oversizedPath }, existingOutput));
            Assert.Equal(sentinel, File.ReadAllBytes(existingOutput));
            Assert.Throws<PdfReadLimitException>(() =>
                PdfPageImporter.AppendPages(validPath, oversizedPath, existingOutput, 1));
            Assert.Equal(sentinel, File.ReadAllBytes(existingOutput));

            Assert.Throws<PdfReadLimitException>(() =>
                PdfMerger.MergeFiles(Path.Combine(absentOutputDirectory, "merged.pdf"), validPath, oversizedPath));
            Assert.False(Directory.Exists(absentOutputDirectory));
            Assert.Throws<PdfReadLimitException>(() =>
                PdfPageImporter.AppendPages(validPath, oversizedPath, Path.Combine(absentOutputDirectory, "imported.pdf"), 1));
            Assert.False(Directory.Exists(absentOutputDirectory));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void MergeAndImportStreamsContinueFromCallerPosition() {
        byte[] first = CreatePdf("first");
        byte[] second = CreatePdf("second");
        using var firstStream = CreatePrefixedStream(first);
        using var secondStream = CreatePrefixedStream(second);

        byte[] merged = PdfMerger.Merge((IEnumerable<Stream>)new[] { firstStream, secondStream });
        Assert.Equal(2, PdfInspector.Inspect(merged).PageCount);

        using var targetStream = CreatePrefixedStream(first);
        using var sourceStream = CreatePrefixedStream(second);
        byte[] imported = PdfPageImporter.AppendPages(targetStream, sourceStream);
        Assert.Equal(2, PdfInspector.Inspect(imported).PageCount);
    }

    private static byte[] CreatePdf(string text) =>
        PdfDocument.Create().Paragraph(paragraph => paragraph.Text(text)).ToBytes();

    private static MemoryStream CreatePrefixedStream(byte[] pdf) {
        var stream = new MemoryStream();
        stream.Write(new byte[] { 9, 8, 7 }, 0, 3);
        stream.Write(pdf, 0, pdf.Length);
        stream.Position = 3;
        return stream;
    }

    private static OversizedLengthStream CreateOversizedStream() {
        var stream = new OversizedLengthStream(PdfLoadOptions.Default.Limits.MaxInputBytes + 4) { Position = 3 };
        return stream;
    }

    private sealed class OversizedLengthStream : Stream {
        private readonly long _length;
        private long _position;

        internal OversizedLengthStream(long length) => _length = length;
        internal bool WasRead { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _length;
        public override long Position { get => _position; set => _position = value; }
        public override int Read(byte[] buffer, int offset, int count) {
            WasRead = true;
            throw new InvalidOperationException("Oversized input should be rejected before reading.");
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
