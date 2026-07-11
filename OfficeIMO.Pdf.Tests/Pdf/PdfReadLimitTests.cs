using System.Diagnostics;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfReadLimitTests {
    [Fact]
    public void InputByteBudgetStopsBeforeObjectScanning() {
        byte[] pdf = BuildPdf();
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = 16 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Load(pdf, options));

        Assert.Equal(PdfReadLimitKind.InputBytes, exception.Kind);
        Assert.Equal(16, exception.Limit);
        Assert.Equal(pdf.Length, exception.Actual);
    }

    [Fact]
    public void SeekableStreamBudgetStopsBeforeBuffering() {
        byte[] pdf = BuildPdf();
        using var stream = new MemoryStream(pdf);
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = 16 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Load(stream, options));

        Assert.Equal(PdfReadLimitKind.InputBytes, exception.Kind);
        Assert.Equal(0, stream.Position);
    }

    [Fact]
    public void IndirectObjectBudgetStopsExcessiveDeclarations() {
        byte[] pdf = BuildPdf();
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxIndirectObjects = 1 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Load(pdf, options));

        Assert.Equal(PdfReadLimitKind.IndirectObjects, exception.Kind);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void RawStreamBudgetStopsAllocation() {
        byte[] pdf = BuildPdf();
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxRawStreamBytes = 4 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Load(pdf, options));

        Assert.Equal(PdfReadLimitKind.RawStreamBytes, exception.Kind);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void InvalidReadBudgetsAreRejectedExplicitly() {
        byte[] pdf = BuildPdf();
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxIndirectObjects = 0 }
        };

        Assert.Throws<ArgumentOutOfRangeException>(() => PdfReadDocument.Load(pdf, options));
    }

    [Fact]
    public void DeterministicHostileInputMutationsRemainBounded() {
        byte[] source = BuildPdf();
        var random = new Random(0x2062);
        var timer = Stopwatch.StartNew();
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits {
                MaxInputBytes = 2 * 1024 * 1024,
                MaxIndirectObjects = 2_000,
                MaxRawStreamBytes = 512 * 1024,
                MaxObjectParsingTime = TimeSpan.FromSeconds(1)
            }
        };

        for (int caseNumber = 0; caseNumber < 32; caseNumber++) {
            int length = random.Next(1, source.Length + 65);
            var candidate = new byte[length];
            Buffer.BlockCopy(source, 0, candidate, 0, Math.Min(source.Length, candidate.Length));
            for (int mutation = 0; mutation < 8; mutation++) {
                candidate[random.Next(candidate.Length)] = (byte)random.Next(256);
            }

            try {
                _ = PdfReadDocument.Load(candidate, options);
            } catch (Exception exception) when (
                exception is ArgumentException ||
                exception is FormatException ||
                exception is InvalidOperationException ||
                exception is IOException) {
                // Malformed candidates may fail, but must stay within the declared parser contract.
            }
        }

        Assert.True(timer.Elapsed < TimeSpan.FromSeconds(10), "Hostile-input parser pass exceeded the test budget: " + timer.Elapsed + ".");
    }

    private static byte[] BuildPdf() => PdfDocument.Create()
        .Paragraph(paragraph => paragraph.Text("Bounded parser source"))
        .ToBytes();
}
