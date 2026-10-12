using System.Text.RegularExpressions;

namespace OfficeIMO.DjVu.Tests;

public sealed class Iw44DecoderTests {
    [Fact]
    public void ProgressiveColorChunksRecoverIndependentReferencePixels() {
        string folder = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Authored");
        var document = DjVuDocument.Load(Path.Combine(folder, "gradient.djvu"));
        var page = Assert.Single(document.Pages);
        var decoder = new Iw44Decoder(new DjVuReadBudget(new DjVuReadOptions(), default));
        var chunks = document.PageChunks(page.Component, default).Where(c => c.Id == "BG44").ToArray();
        Assert.True(chunks.Length > 1);
        foreach (var chunk in chunks) decoder.Read(chunk);
        byte[] expectedFile = File.ReadAllBytes(Path.Combine(folder, "gradient-reference.ppm"));
        var header = Regex.Match(Encoding.ASCII.GetString(expectedFile, 0, 64), @"^P6\s+(\d+)\s+(\d+)\s+255\s");
        Assert.True(header.Success);
        Assert.Equal((int.Parse(header.Groups[1].Value), int.Parse(header.Groups[2].Value)), (decoder.Width, decoder.Height));
        byte[] expected = expectedFile.Skip(header.Length).ToArray(), actual = decoder.Reconstruct();
        Assert.Equal(expected.Length, actual.Length);
        int differences = 0, first = -1, maximum = 0;
        long total = 0;
        for (int i = 0; i < actual.Length; i++) {
            int difference = Math.Abs(actual[i] - expected[i]);
            if (difference != 0) { differences++; if (first == -1) first = i; }
            maximum = Math.Max(maximum, difference); total += difference;
        }
        Assert.True(differences == 0, $"{differences} differing samples; first {first}: expected {expected[Math.Max(first, 0)]}, actual {actual[Math.Max(first, 0)]}; max {maximum}; mean {(double)total / actual.Length:F4}.");
    }
}
