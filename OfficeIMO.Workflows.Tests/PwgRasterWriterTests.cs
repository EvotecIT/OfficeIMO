using System.Diagnostics;

namespace OfficeIMO.Workflows.Tests;

public sealed class PwgRasterWriterTests {
    internal const string RasterToTiff = "/usr/libexec/cups/filter/rastertotiff";
    internal const string IppEvePcl = "/usr/libexec/cups/command/ippevepcl";

    [Fact]
    public void CompressesRepeatedLinesAndPixelRuns() {
        var pixels = new byte[4 * 3];
        for (int i = 0; i < pixels.Length; i++) pixels[i] = 255;
        pixels[4 * 2 + 1] = 0; // third row: white, black, white, white

        byte[] raster = Write(new PwgRasterPage(4, 3, 300, 300, PwgRasterColorSpace.Gray8, pixels));

        Assert.Equal(new byte[] { (byte)'R', (byte)'a', (byte)'S', (byte)'2' }, raster.Take(4));
        // Rows 1-2 repeat (count 1) as a run of four white pixels; row 3 is two literal pixels then a run of two.
        Assert.Equal(new byte[] { 1, 3, 255, 0, 255, 255, 0, 1, 255 }, raster.Skip(4 + PwgRasterWriter.HeaderSize));
    }

    [ExternalToolFact(RasterToTiff)]
    public void CupsRasterReaderDecodesIdenticalRgbPixels() {
        PwgRasterPage page = Pattern(PwgRasterColorSpace.Rgb8, 161, 300, seed: 2);
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pwg-" + Guid.NewGuid().ToString("N") + ".pwg");
        try {
            File.WriteAllBytes(path, Write(page));
            // rastertotiff reads the stream with the CUPS raster reader and writes uncompressed TIFF samples.
            (int exit, byte[] tiff, string error) = Run(RasterToTiff, new[] { "1", "user", "title", "1", "", path }, null);
            Assert.True(exit == 0 && tiff.Length > 0, error);
            Assert.Equal(page.Pixels, UncompressedTiffSamples(tiff));
        } finally {
            File.Delete(path);
        }
    }

    [ExternalToolFact(IppEvePcl)]
    public void IppEvePclConverterAcceptsTheRaster() {
        byte[] raster = Write(Pattern(PwgRasterColorSpace.Gray8, 600, 400, seed: 3));
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pwg-" + Guid.NewGuid().ToString("N") + ".pwg");
        try {
            File.WriteAllBytes(path, raster);
            (int exit, byte[] output, string error) = Run(IppEvePcl, new[] { path }, new Dictionary<string, string> { ["CONTENT_TYPE"] = "image/pwg-raster" });
            Assert.True(exit == 0, error);
            Assert.True(output.Length > 1000);
        } finally {
            File.Delete(path);
        }
    }

    internal static PwgRasterPage Pattern(PwgRasterColorSpace colorSpace, int width, int height, int seed) {
        int bpp = colorSpace == PwgRasterColorSpace.Rgb8 ? 3 : 1;
        var pixels = new byte[width * height * bpp];
        var random = new Random(seed);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                for (int c = 0; c < bpp; c++) {
                    int i = (y * width + x) * bpp + c;
                    // Mix blank bands (line repeats), solid runs, gradients and noise (literal packets).
                    pixels[i] = (y / 16 % 4) switch {
                        0 => 255,
                        1 => (byte)(x / 7 * 31 + c * 50),
                        2 => (byte)random.Next(256),
                        _ => (byte)(x < width / 2 ? 0 : x + y)
                    };
                }
            }
        }
        return new PwgRasterPage(width, height, 300, 300, colorSpace, pixels) { PageSizeName = "custom" };
    }

    internal static byte[] Write(params PwgRasterPage[] pages) {
        using var output = new MemoryStream();
        var writer = new PwgRasterWriter(output, pages.Length);
        foreach (PwgRasterPage page in pages) writer.WritePage(page);
        return output.ToArray();
    }

    /// <summary>Concatenates the strips of a big-endian, uncompressed TIFF written by CUPS.</summary>
    private static byte[] UncompressedTiffSamples(byte[] tiff) {
        Assert.Equal((byte)'M', tiff[0]);
        int ifd = UInt(tiff, 4);
        int entries = UShort(tiff, ifd);
        int[] offsets = Array.Empty<int>(), counts = Array.Empty<int>();
        for (int i = 0; i < entries; i++) {
            int entry = ifd + 2 + i * 12;
            int tag = UShort(tiff, entry);
            if (tag == 259) Assert.Equal(1, UShort(tiff, entry + 8)); // no compression
            if (tag is 273 or 279) {
                int type = UShort(tiff, entry + 2), count = UInt(tiff, entry + 4), size = type == 3 ? 2 : 4;
                int at = count * size <= 4 ? entry + 8 : UInt(tiff, entry + 8);
                int[] values = Enumerable.Range(0, count).Select(n => size == 2 ? UShort(tiff, at + n * 2) : UInt(tiff, at + n * 4)).ToArray();
                if (tag == 273) offsets = values; else counts = values;
            }
        }
        Assert.Equal(offsets.Length, counts.Length);
        return offsets.Zip(counts, (offset, count) => tiff.Skip(offset).Take(count)).SelectMany(strip => strip).ToArray();
    }

    private static int UShort(byte[] data, int offset) => (data[offset] << 8) | data[offset + 1];

    private static int UInt(byte[] data, int offset) => (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];

    internal static (int Exit, byte[] Output, string Error) Run(string tool, IEnumerable<string> arguments, IDictionary<string, string>? environment) {
        var start = new ProcessStartInfo(tool) { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
        if (environment is not null) foreach (KeyValuePair<string, string> pair in environment) start.Environment[pair.Key] = pair.Value;
        using Process process = Process.Start(start)!;
        Task<string> error = process.StandardError.ReadToEndAsync();
        using var output = new MemoryStream();
        process.StandardOutput.BaseStream.CopyTo(output);
        Assert.True(process.WaitForExit(60_000), tool + " did not exit.");
        return (process.ExitCode, output.ToArray(), error.Result);
    }
}

/// <summary>Runs only where the named CUPS reference tools are installed (macOS and CUPS hosts).</summary>
public sealed class ExternalToolFactAttribute : FactAttribute {
    public ExternalToolFactAttribute(params string[] paths) {
        string? missing = paths.FirstOrDefault(path => !File.Exists(path));
        if (missing is not null) Skip = $"{missing} is not installed on this host.";
    }
}
