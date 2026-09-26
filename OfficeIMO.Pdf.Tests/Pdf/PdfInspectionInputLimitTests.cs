using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfInspectionInputLimitTests {
    [Fact]
    public void InspectionPathAndStreamRoutesRejectOversizedInputBeforeParsing() {
        byte[] pdf = CreatePdf();
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length - 1 } };
        string path = Path.Combine(Path.GetTempPath(), "officeimo-inspection-limit-" + Guid.NewGuid().ToString("N") + ".pdf");
        try {
            File.WriteAllBytes(path, pdf);
            Action<string>[] pathRoutes = {
                file => PdfInspector.Inspect(file, options),
                file => PdfInspector.InspectPageRanges(file, options, PdfPageRange.From(1, 1)),
                file => PdfInspector.Probe(file, options),
                file => PdfDocument.Preflight(file, options),
                file => PdfValidator.Validate(file, options),
                file => PdfDiagnostics.Analyze(file, options),
                file => PdfDiagnostics.AnalyzeOptimization(file, options),
                file => PdfSignatureValidator.Validate(file, options),
                file => PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfA3B, file, options)
            };
            foreach (Action<string> route in pathRoutes) {
                PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() => route(path));
                Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
                Assert.Equal(pdf.Length - 1, error.Limit);
            }

            byte[] prefixed = new byte[pdf.Length + 3];
            Buffer.BlockCopy(pdf, 0, prefixed, 3, pdf.Length);
            Action<Stream>[] streamRoutes = {
                input => PdfInspector.Inspect(input, options),
                input => PdfInspector.InspectPageRanges(input, options, PdfPageRange.From(1, 1)),
                input => PdfInspector.Probe(input, options),
                input => PdfDocument.Preflight(input, options),
                input => PdfValidator.Validate(input, options),
                input => PdfDiagnostics.Analyze(input, options),
                input => PdfDiagnostics.AnalyzeOptimization(input, options),
                input => PdfSignatureValidator.Validate(input, options),
                input => PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfA3B, input, options)
            };
            foreach (Action<Stream> route in streamRoutes) {
                using var input = new MemoryStream(prefixed);
                input.Position = 3;
                PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() => route(input));
                Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
                Assert.Equal(3, input.Position);
            }
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void InspectionReadsFromCurrentPositionWithExactBudget() {
        byte[] pdf = CreatePdf();
        byte[] prefixed = new byte[pdf.Length + 3];
        Buffer.BlockCopy(pdf, 0, prefixed, 3, pdf.Length);
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length } };
        using var input = new MemoryStream(prefixed);
        input.Position = 3;

        Assert.True(PdfDocument.Preflight(input, options).CanRead);
        Assert.Equal(input.Length, input.Position);
    }

    [Fact]
    public void InspectionBoundsNonSeekableInputAndPreservesCallerReadErrors() {
        byte[] pdf = CreatePdf();
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length - 1 } };
        using var input = new ForwardOnlyStream(pdf);

        Assert.Equal(PdfReadLimitKind.InputBytes,
            Assert.Throws<PdfReadLimitException>(() => PdfDocument.Preflight(input, options)).Kind);
        using var invalid = new InvalidDataStream();
        Assert.Equal("caller stream failed",
            Assert.Throws<InvalidDataException>(() => PdfDocument.Preflight(invalid, options)).Message);
    }

    private static byte[] CreatePdf() => PdfDocument.Create()
        .Paragraph(paragraph => paragraph.Text("Inspection input budget"))
        .ToBytes();

    private class ForwardOnlyStream : Stream {
        private readonly MemoryStream _source;

        internal ForwardOnlyStream(byte[] bytes) => _source = new MemoryStream(bytes);
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => _source.Read(buffer, offset, count);
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _source.Dispose(); base.Dispose(disposing); }
    }

    private sealed class InvalidDataStream : ForwardOnlyStream {
        internal InvalidDataStream() : base(Array.Empty<byte>()) { }
        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidDataException("caller stream failed");
    }
}
