using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TiffLegacyJpegTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegLegacy");

    [Fact]
    public void LegacyTablesStrilesAndInterchangeRetainSourcePixels() {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            byte[] legacy = File.ReadAllBytes(Path.Combine(Corpus, fields[0]));
            byte[] reference = File.ReadAllBytes(Path.Combine(Corpus, Path.GetFileNameWithoutExtension(fields[0]) + ".reference.tif"));
            Assert.True(OfficeImageReader.TryValidateContent(legacy, fields[0], out _), fields[0] + " validation");
            Assert.True(OfficeTiffCodec.TryDecode(legacy, out var actual), fields[0] + " decode");
            Assert.True(OfficeTiffCodec.TryDecode(reference, out var expected), fields[0] + " reference");
            Assert.Equal((int.Parse(fields[1]), int.Parse(fields[2])), (actual!.Width, actual.Height));
            AssertPixelsEqual(expected!, actual, fields[0]);
        }
    }

    [Theory]
    [InlineData("ojpeg_zackthecat_subsamp22_single_strip", 20)]
    [InlineData("ojpeg_chewey_subsamp21_multi_strip", 14)]
    [InlineData("ojpeg_single_strip_no_rowsperstrip_sanitized", 20)]
    public void UpstreamLegacyImagesMatchReconstructionAndBoundedNativeRendering(string name, int maximumDifference) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, "upstream", name + ".tiff"));
        Assert.True(OfficeImageReader.TryValidateContent(bytes, name, out _));
        Assert.True(OfficeTiffCodec.TryDecode(bytes, out var actual));
        Assert.True(OfficeTiffCodec.TryDecode(File.ReadAllBytes(Path.Combine(Corpus, "upstream", name + ".reference.tif")), out var reference));
        AssertPixelsEqual(reference!, actual!, name);
        Assert.True(OfficePngReader.TryDecode(File.ReadAllBytes(Path.Combine(Corpus, "upstream", name + ".png")), out var native));
        Assert.Equal((native!.Width, native.Height), (actual!.Width, actual.Height));
        for (int y = 0; y < actual.Height; y++) for (int x = 0; x < actual.Width; x++) {
            var a = actual.GetPixel(x, y); var n = native.GetPixel(x, y);
            Assert.True(Math.Abs(a.R - n.R) <= maximumDifference && Math.Abs(a.G - n.G) <= maximumDifference &&
                Math.Abs(a.B - n.B) <= maximumDifference && a.A == n.A, $"{name} at {x},{y}");
        }
    }

    [Fact]
    public void InvalidUpstreamDirectoryRecordsRemainRejected() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, "upstream", "ojpeg_single_strip_no_rowsperstrip.tiff"));
        Assert.False(OfficeImageReader.TryValidateContent(bytes, "source.tiff", out _));
        Assert.False(OfficeTiffCodec.TryDecode(bytes, out _));
    }

    [Theory]
    [InlineData(519)]
    [InlineData(520)]
    [InlineData(521)]
    public void InvalidLegacyTablePointersAreRejected(int tag) {
        byte[] bytes = RawBaseline();
        SetValue(bytes, tag, bytes.Length - 1);
        AssertRejected(bytes);
    }

    [Theory]
    [InlineData(512, 2)]
    [InlineData(515, 65536)]
    [InlineData(279, 0)]
    public void UnsupportedProcessRestartAndEmptyEntropyAreRejected(int tag, int value) {
        byte[] bytes = RawBaseline();
        SetValue(bytes, tag, value);
        AssertRejected(bytes);
    }

    [Theory]
    [InlineData(517, 0)]
    [InlineData(517, 8)]
    [InlineData(518, 8)]
    [InlineData(512, 1)]
    public void InvalidLosslessPredictionOrProcessIsRejected(int tag, int value) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, "tiny-lossless-8-le.tif"));
        SetValue(bytes, tag, value);
        AssertRejected(bytes);
    }

    [Fact]
    public void OversubscribedLegacyHuffmanTreeIsRejected() {
        byte[] bytes = RawBaseline();
        int pointer = Value(bytes, 520);
        bytes[pointer] = 3; // Three one-bit codes cannot form a prefix tree.
        AssertRejected(bytes);
    }

    [Theory]
    [InlineData(513)]
    [InlineData(514)]
    [InlineData(256)]
    public void InterchangeBoundsAndFrameGeometryAreChecked(int tag) {
        string path = Directory.GetFiles(Corpus, "TiffJpegLossless-interchange-p1-*.tif").Single(p => !p.Contains(".reference."));
        byte[] bytes = File.ReadAllBytes(path);
        SetValue(bytes, tag, tag == 256 ? Value(bytes, tag) + 1 : bytes.Length + 1);
        AssertRejected(bytes);
    }

    [Fact]
    public void LeadingScanHeaderMustMatchLegacyMetadata() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, "upstream", "ojpeg_chewey_subsamp21_multi_strip.tiff"));
        int scan = Value(bytes, 273);
        Assert.Equal(218, bytes[scan + 1]);
        bytes[scan + 11] = 1; // Baseline spectral selection must start at zero.
        AssertRejected(bytes);
    }

    [Theory]
    [InlineData(8, "le")]
    [InlineData(8, "be")]
    [InlineData(16, "le")]
    [InlineData(16, "be")]
    public void SingleByteLosslessEntropyProducesTheInitialPredictor(int precision, string order) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, $"tiny-lossless-{precision}-{order}.tif"));
        Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image));
        var pixel = image!.GetPixel(0, 0);
        Assert.Equal((128, 128, 128, 255), ((int)pixel.R, pixel.G, pixel.B, pixel.A));
    }

    [Fact]
    public void LegacyReconstructionHonorsCancellationAndMemoryBudgets() {
        byte[] bytes = RawBaseline();
        Assert.True(OfficeTiffCodec.TryDecodePage(bytes, 0, new(), out _));
        Assert.False(OfficeTiffCodec.TryDecodePage(bytes, 0, new() {
            RetainedManagedBytes = OfficeRasterGuards.MaximumDecodedBytes - bytes.Length
        }, out _));
        Assert.False(OfficeTiffCodec.TryValidateAllPages(bytes, new(), 4095));
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeTiffCodec.TryDecodePage(bytes, 0,
            new() { CancellationToken = cancellation.Token }, out _));
    }

    private static byte[] RawBaseline() => File.ReadAllBytes(Path.Combine(Corpus, "raw-p0-be0-t0-q0-pl1-s1.tif"));

    private static void AssertRejected(byte[] bytes) {
        Assert.False(OfficeTiffCodec.TryDecode(bytes, out _));
        Assert.False(OfficeImageReader.TryValidateContent(bytes, "source.tif", out _));
    }

    // Mutation inputs use little-endian fixtures; follow the field's type/count
    // so inline scalars and indirect pointer arrays exercise the same contract.
    private static (int Offset, int Size) Field(byte[] bytes, int tag) {
        Assert.Equal((byte)'I', bytes[0]);
        int directory = BitConverter.ToInt32(bytes, 4);
        int count = BitConverter.ToUInt16(bytes, directory);
        for (int i = 0; i < count; i++) {
            int entry = directory + 2 + 12 * i;
            if (BitConverter.ToUInt16(bytes, entry) != tag) continue;
            int size = BitConverter.ToUInt16(bytes, entry + 2) == 3 ? 2 : 4;
            int values = BitConverter.ToInt32(bytes, entry + 4);
            return (size * values <= 4 ? entry + 8 : BitConverter.ToInt32(bytes, entry + 8), size);
        }
        throw new InvalidOperationException($"Missing TIFF tag {tag}.");
    }

    private static int Value(byte[] bytes, int tag) {
        var field = Field(bytes, tag);
        return field.Size == 2 ? BitConverter.ToUInt16(bytes, field.Offset) : BitConverter.ToInt32(bytes, field.Offset);
    }

    private static void SetValue(byte[] bytes, int tag, int value) {
        var field = Field(bytes, tag);
        // A SHORT overflow is tested through a LONG scalar, which TIFF permits.
        if (field.Size == 2 && value > ushort.MaxValue) {
            int directory = BitConverter.ToInt32(bytes, 4);
            for (int i = 0; i < BitConverter.ToUInt16(bytes, directory); i++) {
                int entry = directory + 2 + 12 * i;
                if (BitConverter.ToUInt16(bytes, entry) == tag) { BitConverter.GetBytes((ushort)4).CopyTo(bytes, entry + 2); break; }
            }
        }
        byte[] encoded = field.Size == 2 && value <= ushort.MaxValue ? BitConverter.GetBytes((ushort)value) : BitConverter.GetBytes(value);
        encoded.CopyTo(bytes, field.Offset);
    }

    private static void AssertPixelsEqual(OfficeRasterImage expected, OfficeRasterImage actual, string file) {
        Assert.Equal((expected.Width, expected.Height), (actual.Width, actual.Height));
        for (int y = 0; y < expected.Height; y++) for (int x = 0; x < expected.Width; x++)
            Assert.True(expected.GetPixel(x, y).Equals(actual.GetPixel(x, y)), $"{file} at {x},{y}");
    }
}
