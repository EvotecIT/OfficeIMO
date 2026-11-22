using System.Buffers;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDocumentSourceLifetimeTests {
    public enum SourceKind {
        ByteArray,
        SeekableStream,
        NonSeekableStream
    }

    [Theory]
    [InlineData(SourceKind.ByteArray)]
    [InlineData(SourceKind.SeekableStream)]
    [InlineData(SourceKind.NonSeekableStream)]
    public void EverySourceKindReadsTheSamePages(SourceKind kind) {
        byte[] pdf = CreateTextPdf(pageCount: 12);

        using PdfDocument document = Load(kind, pdf, options: null);
        PdfDocumentReadResult read = document.Read();

        Assert.Equal(12, read.PageCount);
        Assert.Contains("Marker page 7.", document.Pages.Split()[6].Read().Text, StringComparison.Ordinal);
        Assert.Contains("Marker page 12.", read.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ByteArraySourceIsReadInPlace() {
        byte[] pdf = CreateTextPdf(pageCount: 2);

        using PdfDocument document = PdfDocument.Load(pdf);

        Assert.Same(pdf, document.GetBytesForOperation());
        Assert.Equal(pdf, document.ToBytes());
        Assert.NotSame(pdf, document.ToBytes());
    }

    [Fact]
    public void BorrowedPoolSliceRemainsTheSourceUntilSaveAndReadComplete() {
        byte[] pdf = CreateTextPdf(pageCount: 5);
        using var hash = SHA256.Create();
        string expectedDigest = BitConverter.ToString(hash.ComputeHash(pdf)).Replace("-", string.Empty).ToLowerInvariant();
        PdfArtifactInfo evidence;

        using (var owner = new PooledPdfInput(pdf, prefixLength: 4096)) {
            ReadOnlyMemory<byte> input = owner.Memory.Slice(4096, pdf.Length);
            Assert.Throws<PdfReadLimitException>(() => PdfDocument.Load(input,
                new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length - 1 } }));

            using PdfDocument document = PdfDocument.Load(input,
                new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length } });
            evidence = Assert.IsType<PdfArtifactInfo>(document.Pipeline.Output);
            Assert.Equal(pdf.Length, evidence.ByteCount);

            int readsAfterLoad = owner.SpanRequests;
            Assert.Contains("Marker page 5.", document.Read().Text, StringComparison.Ordinal);
            Assert.True(owner.SpanRequests > readsAfterLoad);

            using var output = new MemoryStream();
            document.Save(output);
            Assert.Equal(pdf, output.ToArray());
        }

        Assert.Equal(expectedDigest, evidence.Sha256);
    }

    [Fact]
    public void MutationPreflightCanReadRevisionEvidenceAfterReturningBorrowedMemory() {
        byte[] pdf = PdfIncrementalUpdater.UpdateMetadata(CreateTextPdf(pageCount: 2), title: "Updated title");
        PdfMutationPlan plan;
        using (var owner = new PooledPdfInput(pdf, prefixLength: 4096)) {
            using PdfDocument document = PdfDocument.Load(owner.Memory.Slice(4096, pdf.Length));
            plan = document.PlanMutation(PdfMutationOperation.UpdateMetadata);
            Assert.True(plan.CanExecute);
            Assert.True(plan.Preflight.Probe.Security.HasPreviousRevision);
        }

        Assert.True(plan.Preflight.CanNormalizeRedactionObjectGraph);
    }

    [Fact]
    public void RedactionSearchKeepsRevisionEvidenceAfterTheCallerReusesArrayMemory() {
        byte[] pdf = PdfIncrementalUpdater.UpdateMetadata(CreateTextPdf(pageCount: 2), title: "Updated title");
        PdfRedactionPlan search;
        using (PdfDocument document = PdfDocument.Load((ReadOnlyMemory<byte>)pdf)) {
            search = document.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("missing phrase"));
            Assert.Empty(search.Matches);
            Assert.True(search.Preflight.Probe.Security.HasPreviousRevision);
        }

        Array.Clear(pdf, 0, pdf.Length);
        Assert.True(search.Preflight.CanNormalizeRedactionObjectGraph);
    }

    [Fact]
    public void StreamSourceIsOwnedSoTheStreamCanBeDisposedAfterLoading() {
        byte[] pdf = CreateTextPdf(pageCount: 3);
        var stream = new MemoryStream(pdf, writable: false);

        using PdfDocument document = PdfDocument.Load(stream);
        stream.Dispose();

        Assert.Contains("Marker page 3.", document.Read().Text, StringComparison.Ordinal);
        Assert.Equal(pdf, document.ToBytes());
    }

    [Theory]
    [InlineData(SourceKind.ByteArray)]
    [InlineData(SourceKind.SeekableStream)]
    [InlineData(SourceKind.NonSeekableStream)]
    public void Aes256ProtectedSourceOpensFromEverySourceKind(SourceKind kind) {
        byte[] pdf = CreateProtectedPdf();
        Assert.Contains("/V 5 /R 6", PdfEncoding.Latin1GetString(pdf), StringComparison.Ordinal);

        using PdfDocument document = Load(kind, pdf, new PdfLoadOptions { Password = "open" });

        PdfDocumentReadResult read = document.Read();
        Assert.Equal(2, read.PageCount);
        Assert.Contains("Protected second page", read.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalEditOfByteArraySourceAppendsToTheOriginalBytesAndLeavesTheArrayAlone() {
        // Protected documents are edited append-only, so the output must start with the caller's bytes.
        byte[] pdf = CreateProtectedPdf();
        byte[] original = (byte[])pdf.Clone();

        using PdfDocument document = PdfDocument.Load(pdf, new PdfLoadOptions { Password = "owner" });
        PdfOperationResult<PdfDocument> rotation = document.Pages.RotateResult(90, "2");
        byte[] rotated = rotation.RequireValue().ToBytes();

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, rotation.MutationPlan!.ExecutionMode);
        Assert.True(rotated.Length > original.Length);
        Assert.True(rotated.AsSpan(0, original.Length).SequenceEqual(original));
        Assert.Equal(original, pdf);
        Assert.Equal(90, PdfReadDocument.Open(rotated, new PdfLoadOptions { Password = "open" }).Pages[1].GetRotationDegrees());
    }

    [Fact]
    public void LeafPagesAreBuiltOnlyWhenThePageIsRead() {
        byte[] pdf = CreateTextPdf(pageCount: 5);

        PdfReadDocument document = PdfReadDocument.OpenOwned(pdf, options: null);

        Assert.Equal(5, document.Pages.Count);
        Assert.All(document.Pages, page => Assert.False(document.Objects[page.ObjectNumber].IsValueBuilt));
        Assert.Contains("Marker page 4.", document.Pages[3].ExtractText(), StringComparison.Ordinal);
        Assert.True(document.Objects[document.Pages[3].ObjectNumber].IsValueBuilt);
        Assert.False(document.Objects[document.Pages[0].ObjectNumber].IsValueBuilt);
        Assert.False(document.RepairReport.HasRepairs);
    }

    [Fact]
    public void LeafPageWithAWrongParentIsStillRepairedWhenItsDictionaryIsDeferred() {
        // Page 4 names the catalog as its parent; the page-tree walk must still report and restore /Parent.
        byte[] pdf = BuildClassicPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 5 0 R >>",
            "<< /Type /Page /Parent 1 0 R /MediaBox [0 0 200 200] /Contents 5 0 R >>",
            "<< /Length 0 >>\nstream\n\nendstream");

        PdfReadDocument document = PdfReadDocument.OpenOwned(pdf, options: null);

        Assert.Equal(2, document.Pages.Count);
        PdfRepairDiagnostic repair = Assert.Single(document.RepairReport.Diagnostics, diagnostic => diagnostic.Code == "broken-page-parent");
        Assert.Equal(4, repair.ObjectNumber);
        var restored = Assert.IsType<PdfReference>(document.Pages[1].PageDictionary.Items["Parent"]);
        Assert.Equal(2, restored.ObjectNumber);
        Assert.False(document.Objects[3].IsValueBuilt);
        Assert.Throws<PdfParseException>(() => PdfReadDocument.OpenOwned(pdf, new PdfLoadOptions { ParsingMode = PdfParsingMode.Strict }));
    }

    private static PdfDocument Load(SourceKind kind, byte[] pdf, PdfLoadOptions? options) => kind switch {
        SourceKind.ByteArray => PdfDocument.Load(pdf, options),
        SourceKind.SeekableStream => PdfDocument.Load(new MemoryStream(pdf, writable: false), options),
        SourceKind.NonSeekableStream => PdfDocument.Load(new ForwardOnlyStream(pdf), options),
        _ => throw new ArgumentOutOfRangeException(nameof(kind))
    };

    private static byte[] CreateProtectedPdf() =>
        PdfDocument.Create(new PdfCreateOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.AddText("Protected first page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.AddText("Protected second page"))
            .ToBytes();

    internal static byte[] CreateTextPdf(int pageCount) {
        PdfDocument document = PdfDocument.Create();
        for (int page = 1; page <= pageCount; page++) {
            if (page > 1) document.PageBreak();
            int number = page;
            document.Paragraph(paragraph => paragraph.AddText("Marker page " + number + "."));
        }

        return document.ToBytes();
    }

    private static byte[] BuildClassicPdf(params string[] bodies) {
        var builder = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>();
        for (int index = 0; index < bodies.Length; index++) {
            offsets.Add(builder.Length);
            builder.Append(index + 1).Append(" 0 obj\n").Append(bodies[index]).Append("\nendobj\n");
        }

        int xref = builder.Length;
        builder.Append("xref\n0 ").Append(bodies.Length + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) builder.Append(offset.ToString("D10")).Append(" 00000 n \n");
        builder.Append("trailer\n<< /Size ").Append(bodies.Length + 1).Append(" /Root 1 0 R >>\nstartxref\n")
            .Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }

    private sealed class PooledPdfInput : MemoryManager<byte> {
        private byte[]? _rented;

        internal PooledPdfInput(byte[] pdf, int prefixLength) {
            _rented = ArrayPool<byte>.Shared.Rent(prefixLength + pdf.Length + 8192);
            var buffer = _rented.AsSpan();
            buffer.Fill((byte)'x');
            pdf.AsSpan().CopyTo(buffer.Slice(prefixLength));
        }

        internal int SpanRequests { get; private set; }

        public override Span<byte> GetSpan() {
            SpanRequests++;
            return _rented ?? throw new ObjectDisposedException(nameof(PooledPdfInput));
        }

        public override MemoryHandle Pin(int elementIndex = 0) => throw new NotSupportedException();

        public override void Unpin() {
        }

        protected override void Dispose(bool disposing) {
            byte[]? rented = Interlocked.Exchange(ref _rented, null);
            if (rented is not null) ArrayPool<byte>.Shared.Return(rented);
        }
    }

    private sealed class ForwardOnlyStream : Stream {
        private readonly MemoryStream _inner;

        internal ForwardOnlyStream(byte[] bytes) => _inner = new MemoryStream(bytes, writable: false);

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
