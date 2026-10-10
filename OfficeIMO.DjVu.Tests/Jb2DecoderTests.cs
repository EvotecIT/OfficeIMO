using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Security.Cryptography;

namespace OfficeIMO.DjVu.Tests;

public sealed class Jb2DecoderTests {
    [Fact]
    public void IndependentEncoderSymbolsRecoverExactMaskPixels() {
        string folder = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Authored");
        var document = DjVuDocument.Load(Path.Combine(folder, "symbols.djvu"));
        var page = Assert.Single(document.Pages);
        var decoded = page.DecodeMask(new DjVuReadBudget(new DjVuReadOptions(), default));
        Assert.Equal((128, 96), (decoded.Width, decoded.Height));
        Assert.NotEmpty(decoded.Placements);
        byte[] expected = ReadPbm(Path.Combine(folder, "symbols-reference.pbm"), out int width, out int height);
        Assert.Equal((width, height), (decoded.Width, decoded.Height));
        var actual = new byte[width * height];
        foreach (var placement in decoded.Placements) {
            for (int y = 0; y < placement.Bitmap.Height; y++) {
                for (int x = 0; x < placement.Bitmap.Width; x++) {
                    int px = x + placement.X, py = y + placement.Y;
                    if (placement.Bitmap.At(x, y) != 0 && (uint)px < (uint)width && (uint)py < (uint)height) actual[(height - 1 - py) * width + px] = 1;
                }
            }
        }
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void ArchivalSymbolsAndRefinementsRecoverExactIndependentMask() {
        string folder = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Archive");
        var page = Assert.Single(DjVuDocument.Load(Path.Combine(folder, "time-machine-16.djvu")).Pages);
        var decoded = page.DecodeMask(new DjVuReadBudget(new DjVuReadOptions(), default));
        var pixels = new byte[decoded.Width * decoded.Height];
        foreach (var placement in decoded.Placements) {
            for (int y = 0; y < placement.Bitmap.Height; y++) for (int x = 0; x < placement.Bitmap.Width; x++) {
                int px = x + placement.X, py = y + placement.Y;
                if (placement.Bitmap.At(x, y) != 0 && (uint)px < (uint)decoded.Width && (uint)py < (uint)decoded.Height)
                    pixels[(decoded.Height - 1 - py) * decoded.Width + px] = 1;
            }
        }
        using var algorithm = SHA256.Create();
        string hash = BitConverter.ToString(algorithm.ComputeHash(pixels)).Replace("-", string.Empty).ToLowerInvariant();
        Assert.Equal(File.ReadAllText(Path.Combine(folder, "time-machine-16-mask.sha256")).Trim(), hash);
    }

    private static byte[] ReadPbm(string path, out int width, out int height) {
        byte[] data = File.ReadAllBytes(path);
        string prefix = Encoding.ASCII.GetString(data, 0, Math.Min(data.Length, 64));
        var match = Regex.Match(prefix, @"^P4\s+(\d+)\s+(\d+)\s");
        Assert.True(match.Success);
        width = int.Parse(match.Groups[1].Value); height = int.Parse(match.Groups[2].Value);
        var pixels = new byte[width * height];
        int stride = (width + 7) / 8;
        for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) pixels[y * width + x] = (byte)((data[match.Length + y * stride + x / 8] >> (7 - x % 8)) & 1);
        return pixels;
    }
}
