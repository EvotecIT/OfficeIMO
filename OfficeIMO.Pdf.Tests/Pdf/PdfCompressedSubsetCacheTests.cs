using OfficeIMO.Pdf;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCompressedSubsetCacheTests {
    [Fact]
    public void EquivalentFontSnapshotsShareExactCompressedSubsetAcrossConcurrentDocuments() {
        byte[] fontData = ReadBundledFont();
        var programs = Enumerable.Range(0, 24)
            .Select(_ => PdfTrueTypeFontProgram.Parse((byte[])fontData.Clone(), "Compressed subset test"))
            .ToArray();
        foreach (PdfTrueTypeFontProgram program in programs) {
            RecordText(program, "Concurrent cache: ABC é €");
        }

        byte[] raw = programs[0].BuildSubsetFontFile();
        byte[] expected = PdfFlateEncoder.Compress(raw);
        var compressed = new byte[programs.Length][];
        Parallel.For(0, programs.Length, index => {
            compressed[index] = programs[index].BuildSubsetFontFile(true, out int rawLength);
            Assert.Equal(raw.Length, rawLength);
        });

        Assert.Equal(expected, compressed[0]);
        foreach (byte[] result in compressed) {
            Assert.Same(compressed[0], result);
        }
        Assert.Same(raw, programs[0].BuildSubsetFontFile(false, out int uncompressedLength));
        Assert.Equal(raw.Length, uncompressedLength);
    }

    [Fact]
    public void CompressedSubsetsRemainIsolatedByGlyphUsage() {
        byte[] fontData = ReadBundledFont();
        PdfTrueTypeFontProgram first = PdfTrueTypeFontProgram.Parse(fontData, "First subset");
        PdfTrueTypeFontProgram second = PdfTrueTypeFontProgram.Parse((byte[])fontData.Clone(), "Second subset");
        RecordText(first, "A");
        RecordText(second, "B");

        byte[] firstRaw = first.BuildSubsetFontFile();
        byte[] secondRaw = second.BuildSubsetFontFile();
        byte[] firstCompressed = first.BuildSubsetFontFile(true, out int firstLength);
        byte[] secondCompressed = second.BuildSubsetFontFile(true, out int secondLength);

        Assert.NotEqual(firstRaw, secondRaw);
        Assert.NotSame(firstCompressed, secondCompressed);
        Assert.Equal(PdfFlateEncoder.Compress(firstRaw), firstCompressed);
        Assert.Equal(PdfFlateEncoder.Compress(secondRaw), secondCompressed);
        Assert.Equal(firstRaw.Length, firstLength);
        Assert.Equal(secondRaw.Length, secondLength);
    }

    [Fact]
    public void CanceledSubsetRequestsRejectBothColdAndCachedPayloads() {
        PdfTrueTypeFontProgram program = PdfTrueTypeFontProgram.Parse(ReadBundledFont(), "Canceled subset");
        RecordText(program, "Cold cache cancellation");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => program.BuildSubsetFontFile(true, out _, cancellation.Token));
        _ = program.BuildSubsetFontFile(true, out _);
        Assert.Throws<OperationCanceledException>(() => program.BuildSubsetFontFile(true, out _, cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => program.BuildSubsetFontFile(false, out _, cancellation.Token));
    }

    private static byte[] ReadBundledFont() {
        string? path = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        return File.ReadAllBytes(path!);
    }

    private static void RecordText(PdfTrueTypeFontProgram program, string text) {
        foreach (char scalar in text) {
            Assert.True(program.TryGetGlyphId(scalar, out int glyph));
            program.RecordGlyphUsage(glyph, scalar);
        }
    }
}
