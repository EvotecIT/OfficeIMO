using System.Security.Cryptography;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSeekableDocumentSourceTests {
    [Fact]
    public void AcceptedWindowReadsSelectedPageWithoutReReadingAttachmentAndSavesInOnePass() {
        byte[] pdf = CreatePdfWithAttachment();
        const int prefixLength = 43;
        using var input = OpenWindow(pdf, prefixLength);
        PdfArtifactInfo evidence;
        var limits = new PdfReadLimits { MaxInputBytes = pdf.Length };
        using (PdfDocument document = PdfDocument.LoadSeekable(input, pdf.Length, new PdfLoadOptions { Limits = limits })) {
            Assert.Equal(prefixLength, input.Position);
            long bytesReadForAdmission = input.BytesRead;
            Assert.True(bytesReadForAdmission >= pdf.Length);
            PdfDocumentInfo info = document.Inspect();
            Assert.Equal(3, info.Pages.Count);
            Assert.Single(info.Attachments);
            Assert.Contains("Marker page 3.", document.Pages[2].ExtractText(), StringComparison.Ordinal);
            long bytesReadForPage = input.BytesRead;
            Assert.InRange(bytesReadForPage - bytesReadForAdmission, 1, pdf.Length / 4);
            Assert.Equal(prefixLength, input.Position);

            using var output = new MemoryStream();
            document.Save(output);
            Assert.Equal(pdf, output.ToArray());
            Assert.Equal(pdf.Length, input.BytesRead - bytesReadForPage);
            Assert.Equal(prefixLength, input.Position);
            evidence = Assert.IsType<PdfArtifactInfo>(document.Pipeline.Output);
            Assert.Equal(pdf.Length, evidence.ByteCount);
            Assert.Equal(pdf.Length, input.BytesRead - bytesReadForPage);
        }

        input.Dispose();
        using var digest = SHA256.Create();
        Assert.Equal(BitConverter.ToString(digest.ComputeHash(pdf)).Replace("-", string.Empty).ToLowerInvariant(),
            evidence.Sha256);
    }

    [Fact]
    public void ActiveXrefStreamKeepsPageContentAvailableForDemandReads() {
        byte[] pdf = PdfExternalDocumentCompatibilityTests.BuildXrefStreamPdfWithTrailingStaleDuplicatePage();
        const int prefixLength = 43;
        using var input = OpenWindow(pdf, prefixLength);
        using PdfDocument document = PdfDocument.LoadSeekable(input, pdf.Length);
        PdfReadPage page = Assert.Single(document.Pages);
        long bytesReadBeforeExtraction = input.BytesRead;

        Assert.Equal("Active xref stream page", page.ExtractText().Trim());
        Assert.True(input.BytesRead > bytesReadBeforeExtraction);
        Assert.Equal(prefixLength, input.Position);
    }

    [Fact]
    public void IncrementalInputKeepsCanonicalRevisionEvidence() {
        byte[] inputPdf = PdfIncrementalUpdater.UpdateMetadata(
            PdfDocumentSourceLifetimeTests.CreateTextPdf(3), title: "Revision title");
        const int prefixLength = 43;
        using var input = OpenWindow(inputPdf, prefixLength);
        using PdfDocument document = PdfDocument.LoadSeekable(input, inputPdf.Length);
        Assert.Equal(prefixLength, input.Position);
        using PdfDocument canonical = PdfDocument.Load(inputPdf);
        Assert.Equal(canonical.Inspect().Security.StartXrefOffsets, document.Inspect().Security.StartXrefOffsets);

        Assert.True(document.Inspect().Security.HasPreviousRevision);
        Assert.Equal("Revision title", document.Properties.Title);
        Assert.Contains("Marker page 3.", document.Read().Text, StringComparison.Ordinal);
        Assert.Equal(inputPdf, document.ToBytes());
    }

    [Fact]
    public void SeekableAdmissionEnforcesRevisionBudgetForNormalUpdates() {
        byte[] pdf = PdfIncrementalUpdater.UpdateMetadata(
            PdfDocumentSourceLifetimeTests.CreateTextPdf(3), title: "Revision title");
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxRevisions = 1 } };
        using PdfDocument canonical = PdfDocument.Load(pdf, options);
        PdfReadLimitException expected = Assert.Throws<PdfReadLimitException>(() => canonical.Inspect());
        using var input = OpenWindow(pdf, 43);
        PdfReadLimitException actual = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.LoadSeekable(input, pdf.Length, options));

        Assert.Equal(PdfReadLimitKind.Revisions, actual.Kind);
        Assert.Equal(expected.Limit, actual.Limit);
        Assert.Equal(expected.Actual, actual.Actual);
        Assert.Equal(43, input.Position);
        Assert.True(input.CanRead);
    }

    [Fact]
    public void SeekablePathEnforcesAdmissionWhenFileSharingRequiresASnapshot() {
        byte[] pdf = PdfIncrementalUpdater.UpdateMetadata(
            PdfDocumentSourceLifetimeTests.CreateTextPdf(2), title: "Revision title");
        string path = Path.GetTempFileName();
        try {
            File.WriteAllBytes(path, pdf);
            using var shared = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete);
            var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxRevisions = 1 } };

            PdfReadLimitException actual = Assert.Throws<PdfReadLimitException>(() => PdfDocument.LoadSeekable(path, options));

            Assert.Equal(PdfReadLimitKind.Revisions, actual.Kind);
            Assert.Equal(1, actual.Limit);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task SavingToTheBorrowedStreamPreservesTheOriginalWindow() {
        byte[] pdf = PdfDocumentSourceLifetimeTests.CreateTextPdf(2);
        const int prefixLength = 43;
        byte[] container = new byte[prefixLength + pdf.Length + 61];
        pdf.CopyTo(container, prefixLength);
        byte[] original = (byte[])container.Clone();
        using var input = new MemoryStream(container, writable: true) { Position = prefixLength };

        using PdfDocument document = PdfDocument.LoadSeekable(input, pdf.Length);
        Assert.Throws<ArgumentException>("stream", () => document.Save(input));
        await Assert.ThrowsAsync<ArgumentException>("stream", () => document.SaveAsync(input));

        var converted = new PdfDocumentConversionResult(document, new PdfConversionReport());
        Assert.Throws<ArgumentException>("stream", () => converted.Save(input, failOnLoss: true));
        await Assert.ThrowsAsync<ArgumentException>("stream", () => converted.SaveAsync(input, failOnLoss: true));

        Assert.Equal(prefixLength, input.Position);
        Assert.Equal(original, container);
        Assert.Equal(pdf, document.ToBytes());
    }

    [Fact]
    public void LostSeekableInputReturnsAFailedSaveResultWithoutRecapturingIt() {
        byte[] pdf = PdfDocumentSourceLifetimeTests.CreateTextPdf(2);
        using var input = new MemoryStream(pdf, writable: false);
        using PdfDocument document = PdfDocument.LoadSeekable(input);
        input.Dispose();
        using var output = new MemoryStream();

        PdfSaveResult result = document.TrySaveWithEvidence(output);

        Assert.False(result.Succeeded);
        Assert.IsType<ObjectDisposedException>(result.Exception);
        Assert.True(result.Pipeline.HasFailures);
    }

    [Fact]
    public void CancellingASeekableReadRestoresPositionAndLeavesInputOpen() {
        byte[] pdf = PdfDocumentSourceLifetimeTests.CreateTextPdf(2);
        const int prefixLength = 43;
        using var cancellation = new CancellationTokenSource();
        using var input = OpenWindow(pdf, prefixLength, afterRead: cancellation.Cancel);

        Assert.ThrowsAny<OperationCanceledException>(() =>
            PdfDocument.LoadSeekable(input, pdf.Length, cancellationToken: cancellation.Token));
        Assert.True(input.CanRead);
        Assert.Equal(prefixLength, input.Position);
    }

    [Fact]
    public void PartialReadFailureInvalidatesTheCursorCacheAndPreservesPosition() {
        byte[] bytes = new byte[150_000];
        bytes[0] = 42;
        bytes[65528] = 97;
        const int prefixLength = 43;
        var expected = new IOException("Input read failed.");
        int reads = 0;
        using var source = OpenWindow(bytes, prefixLength, () => {
            if (++reads == 19) throw expected;
        });
        using var input = new PdfSeekableInput(source, prefixLength, bytes.Length, new PdfReadLimits());
        var cursor = new PdfSeekableCursor(input, CancellationToken.None);

        using (PdfSeekableInput.Admission admission = input.BeginAdmission()) {
            Assert.Equal(42, cursor.ReadByte(0));
            Assert.Same(expected, Assert.Throws<IOException>(() => cursor.ReadByte(70_000)));
            Assert.Same(expected, Assert.Throws<IOException>(() => cursor.ReadByte(0)));
            Assert.Same(expected, Assert.Throws<IOException>(admission.ThrowIfReadFailed));
        }

        Assert.Equal(42, cursor.ReadByte(0));
        Assert.Equal(prefixLength, source.Position);
        Assert.True(source.CanRead);
    }

    private static byte[] CreatePdfWithAttachment() {
        byte[] attachment = new byte[2 * 1024 * 1024];
        new Random(1927).NextBytes(attachment);
        var options = new PdfCreateOptions().AddEmbeddedFile("payload.bin", attachment);
        using PdfDocument document = PdfDocument.Create(options);
        for (int page = 1; page <= 3; page++) {
            if (page > 1) document.Content.PageBreak();
            document.AddParagraph($"Marker page {page}.");
        }
        return document.ToBytes();
    }

    private static PartialReadStream OpenWindow(byte[] pdf, int prefixLength, Action? afterRead = null) {
        byte[] container = new byte[prefixLength + pdf.Length + 61];
        pdf.CopyTo(container, prefixLength);
        return new PartialReadStream(container, prefixLength, pdf.Length, afterRead);
    }

    private sealed class PartialReadStream : Stream {
        private readonly MemoryStream _input;
        private readonly long _start;
        private readonly long _end;
        private readonly Action? _afterRead;

        internal PartialReadStream(byte[] bytes, int start, int length, Action? afterRead) {
            _input = new MemoryStream(bytes, writable: false) { Position = start };
            _start = start;
            _end = start + length;
            _afterRead = afterRead;
        }

        internal long BytesRead { get; private set; }
        public override bool CanRead => _input.CanRead;
        public override bool CanSeek => _input.CanSeek;
        public override bool CanWrite => false;
        public override long Length => _input.Length;
        public override long Position { get => _input.Position; set => _input.Position = value; }
        public override long Seek(long offset, SeekOrigin origin) => _input.Seek(offset, origin);
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Flush() => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override int Read(byte[] buffer, int offset, int count) {
            if (Position < _start || count > _end - Position) throw new InvalidDataException("Read outside PDF window.");
            int bytesRead = _input.Read(buffer, offset, Math.Min(count, 4093));
            BytesRead += bytesRead;
            _afterRead?.Invoke();
            return bytesRead;
        }

        protected override void Dispose(bool disposing) {
            if (disposing) _input.Dispose();
            base.Dispose(disposing);
        }
    }
}
