namespace OfficeIMO.Drawing.Benchmarks;

internal sealed record ImageBenchmarkScenario(
    string Id,
    int Width,
    int Height,
    Func<OfficeRasterImage> CreateImage);

/// <summary>Deterministic image shapes used to expose different codec and resampling costs.</summary>
internal static class ImageBenchmarkScenarios {
    internal static readonly ImageBenchmarkScenario Tiny =
        new("Tiny", 16, 16, () => ImageBenchmarkCorpus.CreatePattern(16, 16));
    internal static readonly ImageBenchmarkScenario Screenshot =
        new("Screenshot", 1366, 768, () => CreateScreenshot(1366, 768));
    internal static readonly ImageBenchmarkScenario Text =
        new("Text", 1200, 800, () => CreateText(1200, 800));
    internal static readonly ImageBenchmarkScenario LineArt =
        new("LineArt", 1200, 800, () => CreateLineArt(1200, 800));
    internal static readonly ImageBenchmarkScenario Scan =
        new("Scan", 1200, 1600, () => CreateScan(1200, 1600));
    internal static readonly ImageBenchmarkScenario AlphaGraphic =
        new("AlphaGraphic", 1024, 1024, () => CreateAlphaGraphic(1024, 1024));
    internal static readonly ImageBenchmarkScenario HighEntropy =
        new("HighEntropy", 1024, 1024, () => CreateHighEntropy(1024, 1024));
    internal static readonly ImageBenchmarkScenario Photo =
        new("Photo", 2048, 1619, () => ImageBenchmarkCorpus.Decode(ImageBenchmarkCorpus.Photo.ReadBytes(), "Photo scenario"));
    internal static readonly ImageBenchmarkScenario VeryLarge =
        new("VeryLarge", 4096, 3072, () => CreateHighEntropy(4096, 3072));

    internal static IReadOnlyList<ImageBenchmarkScenario> All { get; } =
        [Tiny, Screenshot, Text, LineArt, Scan, AlphaGraphic, HighEntropy, Photo, VeryLarge];

    internal static IReadOnlyList<string> TimedIds { get; } =
        [Tiny.Id, Screenshot.Id, AlphaGraphic.Id, HighEntropy.Id, Photo.Id];

    internal static IReadOnlyList<string> ResamplingIds { get; } =
        [Photo.Id, Text.Id, LineArt.Id, AlphaGraphic.Id];

    internal static ImageBenchmarkScenario Get(string id) =>
        All.Single(scenario => string.Equals(scenario.Id, id, StringComparison.Ordinal));

    internal static string Fingerprint(OfficeRasterImage image) {
        ulong hash = 14695981039346656037UL;
        for (int sampleY = 0; sampleY < 16; sampleY++) {
            int y = Math.Min(image.Height - 1, sampleY * image.Height / 16);
            for (int sampleX = 0; sampleX < 16; sampleX++) {
                int x = Math.Min(image.Width - 1, sampleX * image.Width / 16);
                OfficeColor color = image.GetPixel(x, y);
                hash = Mix(hash, color.R);
                hash = Mix(hash, color.G);
                hash = Mix(hash, color.B);
                hash = Mix(hash, color.A);
            }
        }
        return hash.ToString("X16");
    }

    private static ulong Mix(ulong hash, byte value) =>
        unchecked((hash ^ value) * 1099511628211UL);

    private static OfficeRasterImage CreateScreenshot(int width, int height) {
        var image = new OfficeRasterImage(width, height, OfficeColor.FromRgb(246, 248, 250));
        int headerHeight = Math.Max(1, height / 12);
        int sidebarWidth = Math.Max(1, width / 6);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                OfficeColor color;
                if (y < headerHeight) {
                    color = OfficeColor.FromRgb(32, 40, 52);
                } else if (x < sidebarWidth) {
                    color = (y / 34) % 2 == 0
                        ? OfficeColor.FromRgb(230, 235, 241)
                        : OfficeColor.FromRgb(238, 242, 246);
                } else if ((y % 72) < 2 || (x % 220) < 2) {
                    color = OfficeColor.FromRgb(205, 212, 220);
                } else if (((x / 180) + (y / 96)) % 7 == 0) {
                    color = OfficeColor.FromRgb(78, 132, 196);
                } else {
                    byte shade = (byte)(246 - ((x * 3 + y * 5) & 3));
                    color = OfficeColor.FromRgb(shade, (byte)(shade + 1), (byte)(shade + 3));
                }
                image.SetPixel(x, y, color);
            }
        }
        return image;
    }

    private static OfficeRasterImage CreateScan(int width, int height) {
        var image = new OfficeRasterImage(width, height);
        for (int y = 0; y < height; y++) {
            bool textRow = y > height / 10 && y < height * 9 / 10 && (y % 38) < 5;
            for (int x = 0; x < width; x++) {
                int paperNoise = ((x * 17 + y * 31 + (x ^ y) * 3) & 7) - 3;
                int value = 246 + paperNoise;
                if (textRow && x > width / 9 && x < width * (7 + (y / 38) % 2) / 9) {
                    value = 38 + ((x + y) & 15);
                } else if (x == width / 14 || x == width * 13 / 14) {
                    value = 188;
                }
                byte channel = (byte)Math.Max(0, Math.Min(255, value));
                image.SetPixel(x, y, OfficeColor.FromRgb(channel, channel, channel));
            }
        }
        return image;
    }

    private static OfficeRasterImage CreateText(int width, int height) {
        var image = new OfficeRasterImage(width, height, OfficeColor.FromRgb(250, 249, 246));
        for (int y = 0; y < height; y++) {
            int lineY = (y - 55) % 35;
            int line = Math.Max(0, (y - 55) / 35);
            for (int x = 0; x < width; x++) {
                if (y < 43 || lineY < 0 || lineY >= 15 || x < 71 || x >= width - 70) continue;
                int glyph = (x - 71) / 17;
                if (glyph >= 48 + (line % 5) * 6) continue;
                int glyphX = (x - 71) % 17;
                bool ink = (glyphX >= 2 && glyphX <= 4) || lineY < 3 || lineY >= 12 ||
                    ((glyph * 3 + lineY * 5) % 13) == glyphX;
                if (ink && glyphX < 14) image.SetPixel(x, y, OfficeColor.FromRgb(28, 34, 42));
            }
        }
        return image;
    }

    private static OfficeRasterImage CreateLineArt(int width, int height) {
        var image = new OfficeRasterImage(width, height, OfficeColor.White);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                bool grid = x % 96 < 2 || y % 80 < 2;
                bool rising = Math.Abs(((x * 3 + y * 5) % 211) - 105) < 2;
                bool falling = Math.Abs(((x * 7 - y * 4 + 4096) % 257) - 128) < 2;
                if (grid || rising || falling) {
                    image.SetPixel(x, y, OfficeColor.FromRgb(18, 28, 44));
                } else if ((x / 160 + y / 120) % 7 == 0 && x % 160 > 24 && y % 120 > 24) {
                    image.SetPixel(x, y, OfficeColor.FromRgb(40, 116, 184));
                }
            }
        }
        return image;
    }

    private static OfficeRasterImage CreateAlphaGraphic(int width, int height) {
        var image = new OfficeRasterImage(width, height, OfficeColor.Transparent);
        double centerX = (width - 1) / 2D;
        double centerY = (height - 1) / 2D;
        double radius = Math.Min(width, height) * 0.46D;
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                double dx = x - centerX;
                double dy = y - centerY;
                double distance = Math.Sqrt(dx * dx + dy * dy);
                if (distance > radius) continue;
                byte alpha = (byte)Math.Max(0D, Math.Min(255D, (radius - distance) * 12D));
                image.SetPixel(
                    x,
                    y,
                    OfficeColor.FromRgba(
                        (byte)((x * 255L) / Math.Max(1, width - 1)),
                        (byte)((y * 255L) / Math.Max(1, height - 1)),
                        (byte)(180 + ((x ^ y) & 63)),
                        alpha));
            }
        }
        return image;
    }

    private static OfficeRasterImage CreateHighEntropy(int width, int height) {
        var image = new OfficeRasterImage(width, height);
        uint state = 0x9E3779B9U;
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                state ^= state << 13;
                state ^= state >> 17;
                state ^= state << 5;
                image.SetPixel(
                    x,
                    y,
                    OfficeColor.FromRgba(
                        (byte)state,
                        (byte)(state >> 8),
                        (byte)(state >> 16),
                        (byte)(96 + ((state >> 24) % 160))));
            }
        }
        return image;
    }
}
