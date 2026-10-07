using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TiffJpegTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpeg");

    [Theory]
    [InlineData("TiffJpeg")]
    [InlineData("TiffJpegExtended")]
    public void IndependentJpegStripsTilesAndTablesPreserveTiffColors(string folder) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", folder);
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            string name = fields[0];
            int photo = int.Parse(fields[1]), samples = photo == 5 ? 4 : photo == 2 || photo == 6 ? 3 : 1;
            byte[] bytes = File.ReadAllBytes(Path.Combine(corpus, name));
            Assert.True(OfficeImageReader.TryValidateContent(bytes, name, out _), name + " validation");
            Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image), name + " decode");
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, name + ".raw"));
            for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                int p = (y * 35 + x) * samples;
                int r, g, b;
                if (photo == 0 || photo == 1) r = g = b = photo == 0 ? 255 - expected[p] : expected[p];
                else if (photo == 5) {
                    r = 255 - Math.Min(255, expected[p] + expected[p + 3]);
                    g = 255 - Math.Min(255, expected[p + 1] + expected[p + 3]);
                    b = 255 - Math.Min(255, expected[p + 2] + expected[p + 3]);
                } else { r = expected[p]; g = expected[p + 1]; b = expected[p + 2]; }
                OfficeColor actual = image!.GetPixel(x, y);
                int error = Math.Max(Math.Abs(actual.R - r), Math.Max(Math.Abs(actual.G - g), Math.Abs(actual.B - b)));
                Assert.True(error <= 3, $"{name} {x},{y}: expected {r},{g},{b}; got {actual.R},{actual.G},{actual.B}; delta {error}");
                Assert.Equal(255, actual.A);
            }
        }
    }
    [Theory]
    [InlineData("wide-dc.jpg")]
    [InlineData("wide-ac.jpg")]
    public void WideQuantizationSaturatesAfterInverseTransform(string name) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegExtended", name));
        Assert.True(OfficeJpegCodec.TryDecode(bytes, out var image));
        // A single DC coefficient of 65535 produces a uniformly white block.
        // A single horizontal AC coefficient produces four white then four black columns.
        for (int y = 0; y < 8; y++) for (int x = 0; x < 8; x++) {
            byte expected = name == "wide-dc.jpg" || x < 4 ? (byte)255 : (byte)0;
            var pixel = image!.GetPixel(x, y);
            Assert.Equal(expected, pixel.R); Assert.Equal(expected, pixel.G); Assert.Equal(expected, pixel.B);
            Assert.Equal(255, pixel.A);
        }
    }

    [Fact]
    public void ExtendedSequentialWideCoefficientClipsAfterInverseTransform() {
        // Eight-bit SOF1, sixteen-bit quantization table (65535), and one AC
        // coefficient of 2 at natural index 4. Its IDCT has alternating signs;
        // clip the final samples rather than overflowing the integer workspace.
        byte[] bytes = Convert.FromBase64String(
            "/9j/2wCDEP///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8EACwgACAAIAQERAP/EACcAAQAAAAAAAAAAAAAAAAAAAAAQAAIAAAAAAAAAAAAAAAAAANIA/9oACAEBAAA/ABP/2Q==");
        Assert.True(OfficeJpegCodec.TryDecode(bytes, out var image));
        Assert.Equal(8, image!.Width);
        Assert.Equal(8, image.Height);
        byte[] expected = { 255, 0, 0, 255, 255, 0, 0, 255 };
        for (int y = 0; y < 8; y++) for (int x = 0; x < 8; x++) {
            var pixel = image.GetPixel(x, y);
            Assert.Equal(expected[x], pixel.R);
            Assert.Equal(expected[x], pixel.G);
            Assert.Equal(expected[x], pixel.B);
            Assert.Equal(255, pixel.A);
        }
    }

    [Theory]
    [InlineData(0)] // Segment dimensions disagree with the TIFF strip.
    [InlineData(1)] // Progressive frames are outside TIFF Technical Note 2.
    [InlineData(2)] // Missing end-of-image marker.
    [InlineData(3)] // Segments mix baseline and extended sequential processes.
    public void MalformedJpegSegmentsFailValidationAndDecode(int fault) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, "p2-be0-t0-q0-pl1-s1.tif"));
        var (offsetSlot, lengthSlot) = FirstStripSlots(bytes);
        int start = BitConverter.ToInt32(bytes, offsetSlot), length = BitConverter.ToUInt16(bytes, lengthSlot);
        int frame = FindMarker(bytes, start, 192);
        if (fault == 0) bytes[frame + 8]++;
        else if (fault == 1) bytes[frame + 1] = 194;
        else if (fault == 2) bytes[start + length - 1] = 0;
        else bytes[frame + 1] = 193;
        Assert.False(OfficeTiffCodec.TryDecode(bytes, out _));
        Assert.False(OfficeImageReader.TryValidateContent(bytes, "malformed.tif", out _));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(5)]
    public void TiffOwnsComponentOrderAndIgnoresAdobeColorOverrides(int photo) {
        byte[] original = File.ReadAllBytes(Path.Combine(Corpus, $"p{photo}-be0-t0-q0-pl1-s1.tif"));
        Assert.True(OfficeTiffCodec.TryDecode(original, out var expected));
        var (offsetSlot, lengthSlot) = FirstStripSlots(original);
        int start = BitConverter.ToInt32(original, offsetSlot), length = BitConverter.ToUInt16(original, lengthSlot);
        byte[] app14 = { 255, 238, 0, 14, 65, 100, 111, 98, 101, 0, 100, 0, 0, 0, 0, 2 };
        var bytes = new byte[original.Length + length + app14.Length];
        original.CopyTo(bytes, 0);
        int appended = original.Length;
        Array.Copy(original, start, bytes, appended, 2);
        app14.CopyTo(bytes, appended + 2);
        Array.Copy(original, start + 2, bytes, appended + 2 + app14.Length, length - 2);
        BitConverter.GetBytes(appended).CopyTo(bytes, offsetSlot);
        BitConverter.GetBytes((ushort)(length + app14.Length)).CopyTo(bytes, lengthSlot);
        int frame = FindMarker(bytes, appended, 192), scan = FindMarker(bytes, appended, 218);
        int samples = photo == 5 ? 4 : 3;
        byte[] ids = photo == 5 ? new byte[] { 77, 89, 75, 67 } : new byte[] { 66, 82, 71 };
        for (int i = 0; i < samples; i++) {
            bytes[frame + 10 + i * 3] = ids[i];
            bytes[scan + 5 + i * 2] = ids[i];
        }
        Assert.True(OfficeImageReader.TryValidateContent(bytes, "metadata.tif", out _));
        Assert.True(OfficeTiffCodec.TryDecode(bytes, out var actual));
        for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++)
            Assert.Equal(expected!.GetPixel(x, y), actual!.GetPixel(x, y));
    }

    private static (int Offset, int Length) FirstStripSlots(byte[] bytes) {
        int ifd = BitConverter.ToInt32(bytes, 4), count = BitConverter.ToUInt16(bytes, ifd);
        int offset = 0, length = 0;
        for (int i = 0; i < count; i++) {
            int entry = ifd + 2 + i * 12, tag = BitConverter.ToUInt16(bytes, entry);
            if (tag == 273) offset = BitConverter.ToInt32(bytes, entry + 8);
            if (tag == 279) {
                Assert.Equal(3, BitConverter.ToUInt16(bytes, entry + 2));
                length = entry + 8;
            }
        }
        Assert.NotEqual(0, offset); Assert.NotEqual(0, length);
        return (offset, length);
    }

    private static int FindMarker(byte[] bytes, int start, int marker) {
        for (int i = start + 2; i + 1 < bytes.Length; i++)
            if (bytes[i] == 255 && bytes[i + 1] == marker) return i;
        throw new InvalidOperationException("Missing fixture JPEG marker.");
    }

}
