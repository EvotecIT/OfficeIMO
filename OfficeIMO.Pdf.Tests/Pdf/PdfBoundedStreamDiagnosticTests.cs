using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfBoundedStreamDiagnosticTests {
    [Fact]
    public void GrowingRemainingStreamReportsObservedInputBeyondLimit() {
        using var stream = new GrowingLengthStream(new byte[] { 1, 2, 3, 4, 5 });
        var options = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = 4 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(
            () => PdfDocumentSource.FromRemainingStream(stream, options));

        Assert.Equal(PdfReadLimitKind.InputBytes, exception.Kind);
        Assert.Equal(4, exception.Limit);
        Assert.Equal(5, exception.Actual);
    }

    private sealed class GrowingLengthStream : MemoryStream {
        private int _lengthReads;

        internal GrowingLengthStream(byte[] bytes) : base(bytes) { }

        public override long Length => ++_lengthReads <= 2 ? base.Length - 1 : base.Length;
    }
}
