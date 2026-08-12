using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Tests {
    internal static class VisualBaselineTestSupport {
        internal static VisualRasterComparison CompareRasterImages(
            byte[] expectedPng,
            byte[] actualPng,
            int channelTolerance,
            int allowedDifferentPixels,
            double maximumMeanAbsoluteError = double.PositiveInfinity,
            double maximumRootMeanSquareError = double.PositiveInfinity,
            double maximumMeanLuminanceError = double.PositiveInfinity) {
            OfficeRasterImage expected = DecodePng(expectedPng, "Expected visual baseline is not a supported PNG file.");
            OfficeRasterImage actual = DecodePng(actualPng, "Actual visual output is not a supported PNG file.");
            return CompareRasterImages(
                expected,
                actual,
                channelTolerance,
                allowedDifferentPixels,
                maximumMeanAbsoluteError,
                maximumRootMeanSquareError,
                maximumMeanLuminanceError);
        }

        internal static VisualRasterComparison CompareRasterImages(
            OfficeRasterImage expected,
            OfficeRasterImage actual,
            int channelTolerance,
            int allowedDifferentPixels,
            double maximumMeanAbsoluteError = double.PositiveInfinity,
            double maximumRootMeanSquareError = double.PositiveInfinity,
            double maximumMeanLuminanceError = double.PositiveInfinity) {
            if (expected.Width != actual.Width || expected.Height != actual.Height) {
                OfficeRasterImage sizeDiff = new OfficeRasterImage(Math.Max(expected.Width, actual.Width), Math.Max(expected.Height, actual.Height), OfficeColor.White);
                OfficeRasterCanvas canvas = new OfficeRasterCanvas(sizeDiff);
                canvas.FillRectangle(0, 0, expected.Width, expected.Height, OfficeColor.FromRgba(37, 99, 235, 180));
                canvas.FillRectangle(0, 0, actual.Width, actual.Height, OfficeColor.FromRgba(220, 38, 38, 180));
                return new VisualRasterComparison(
                    false,
                    0,
                    Math.Max(expected.Width * expected.Height, actual.Width * actual.Height),
                    255,
                    channelTolerance,
                    allowedDifferentPixels,
                    double.PositiveInfinity,
                    maximumMeanAbsoluteError,
                    double.PositiveInfinity,
                    maximumRootMeanSquareError,
                    double.PositiveInfinity,
                    maximumMeanLuminanceError,
                    OfficePngWriter.Encode(sizeDiff));
            }

            int differentPixels = 0;
            int maxChannelDelta = 0;
            long absoluteError = 0;
            long squaredError = 0;
            double luminanceError = 0D;
            OfficeRasterImage diff = new OfficeRasterImage(expected.Width, expected.Height, OfficeColor.White);
            for (int y = 0; y < expected.Height; y++) {
                for (int x = 0; x < expected.Width; x++) {
                    OfficeColor expectedPixel = expected.GetPixel(x, y);
                    OfficeColor actualPixel = actual.GetPixel(x, y);
                    int deltaR = Math.Abs(expectedPixel.R - actualPixel.R);
                    int deltaG = Math.Abs(expectedPixel.G - actualPixel.G);
                    int deltaB = Math.Abs(expectedPixel.B - actualPixel.B);
                    int deltaA = Math.Abs(expectedPixel.A - actualPixel.A);
                    int maxPixelDelta = Math.Max(Math.Max(deltaR, deltaG), Math.Max(deltaB, deltaA));
                    maxChannelDelta = Math.Max(maxChannelDelta, maxPixelDelta);
                    absoluteError += deltaR + deltaG + deltaB + deltaA;
                    squaredError +=
                        deltaR * deltaR +
                        deltaG * deltaG +
                        deltaB * deltaB +
                        deltaA * deltaA;
                    double expectedLuminance =
                        expectedPixel.R * 0.2126D +
                        expectedPixel.G * 0.7152D +
                        expectedPixel.B * 0.0722D;
                    double actualLuminance =
                        actualPixel.R * 0.2126D +
                        actualPixel.G * 0.7152D +
                        actualPixel.B * 0.0722D;
                    luminanceError += Math.Abs(expectedLuminance - actualLuminance);
                    if (maxPixelDelta > channelTolerance) {
                        differentPixels++;
                        diff.SetPixel(x, y, OfficeColor.FromRgb(255, (byte)Math.Min(255, Math.Max(deltaR, deltaG) * 4), (byte)Math.Min(255, Math.Max(deltaB, deltaA) * 4)));
                    } else {
                        int gray = (expectedPixel.R + expectedPixel.G + expectedPixel.B) / 3;
                        byte muted = (byte)(240 - Math.Min(120, gray / 3));
                        diff.SetPixel(x, y, OfficeColor.FromRgb(muted, muted, muted));
                    }
                }
            }

            int totalPixels = expected.Width * expected.Height;
            double totalChannels = totalPixels * 4D;
            double meanAbsoluteError = absoluteError / totalChannels;
            double rootMeanSquareError = Math.Sqrt(squaredError / totalChannels);
            double meanLuminanceError = luminanceError / totalPixels;
            return new VisualRasterComparison(
                differentPixels <= allowedDifferentPixels &&
                meanAbsoluteError <= maximumMeanAbsoluteError &&
                rootMeanSquareError <= maximumRootMeanSquareError &&
                meanLuminanceError <= maximumMeanLuminanceError,
                differentPixels,
                totalPixels,
                maxChannelDelta,
                channelTolerance,
                allowedDifferentPixels,
                meanAbsoluteError,
                maximumMeanAbsoluteError,
                rootMeanSquareError,
                maximumRootMeanSquareError,
                meanLuminanceError,
                maximumMeanLuminanceError,
                OfficePngWriter.Encode(diff));
        }

        internal static OfficeRasterImage DecodePng(byte[] bytes, string failureMessage) {
            ValidatePngContainerIndependently(bytes, failureMessage);
            if (!OfficePngReader.TryDecode(bytes, out OfficeRasterImage? image) || image == null) {
                throw new InvalidOperationException(failureMessage);
            }

            return image;
        }

        private static void ValidatePngContainerIndependently(byte[] bytes, string failureMessage) {
            byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
            if (bytes == null || bytes.Length < 33 || !signature.SequenceEqual(bytes.Take(signature.Length))) {
                throw new InvalidOperationException(failureMessage + " PNG signature is invalid.");
            }

            int offset = signature.Length;
            bool seenHeader = false;
            bool seenData = false;
            while (offset + 12 <= bytes.Length) {
                int length = ReadBigEndianInt32(bytes, offset);
                long end = (long)offset + 12L + length;
                if (length < 0 || end > bytes.Length) {
                    throw new InvalidOperationException(failureMessage + " PNG chunk length is invalid.");
                }
                string type = System.Text.Encoding.ASCII.GetString(bytes, offset + 4, 4);
                if (!seenHeader && (type != "IHDR" || length != 13)) {
                    throw new InvalidOperationException(failureMessage + " PNG does not start with IHDR.");
                }
                uint expected = ReadBigEndianUInt32(bytes, offset + 8 + length);
                uint actual = ComputePngCrc(bytes, offset + 4, 4 + length);
                if (expected != actual) {
                    throw new InvalidOperationException(failureMessage + " PNG chunk CRC is invalid.");
                }
                if (type == "IHDR") {
                    if (seenHeader) throw new InvalidOperationException(failureMessage + " PNG repeats IHDR.");
                    seenHeader = true;
                } else if (type == "IDAT") {
                    seenData = true;
                } else if (type == "IEND") {
                    if (length != 0 || !seenData || end != bytes.Length) {
                        throw new InvalidOperationException(failureMessage + " PNG ending is invalid.");
                    }
                    return;
                }
                offset = (int)end;
            }
            throw new InvalidOperationException(failureMessage + " PNG is incomplete.");
        }

        private static int ReadBigEndianInt32(byte[] bytes, int offset) =>
            (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

        private static uint ReadBigEndianUInt32(byte[] bytes, int offset) =>
            ((uint)bytes[offset] << 24) | ((uint)bytes[offset + 1] << 16) | ((uint)bytes[offset + 2] << 8) | bytes[offset + 3];

        private static uint ComputePngCrc(byte[] bytes, int offset, int count) {
            uint crc = 0xFFFFFFFFU;
            for (int index = 0; index < count; index++) {
                crc ^= bytes[offset + index];
                for (int bit = 0; bit < 8; bit++) {
                    crc = (crc & 1U) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
                }
            }
            return crc ^ 0xFFFFFFFFU;
        }

        internal static byte[] CreateRgbPng(int width, int height, byte[] rgb) {
            if (width <= 0 || height <= 0) {
                throw new ArgumentOutOfRangeException(nameof(width), "PNG dimensions must be positive.");
            }

            if (rgb.Length != width * height * 3) {
                throw new ArgumentException("RGB buffer length does not match PNG dimensions.", nameof(rgb));
            }

            OfficeRasterImage image = new OfficeRasterImage(width, height, OfficeColor.Transparent);
            int source = 0;
            for (int y = 0; y < height; y++) {
                for (int x = 0; x < width; x++) {
                    image.SetPixel(x, y, OfficeColor.FromRgb(rgb[source], rgb[source + 1], rgb[source + 2]));
                    source += 3;
                }
            }

            return OfficePngWriter.Encode(image);
        }

        internal static int CountNonBackgroundPixels(OfficeRasterImage image, OfficeColor background, int channelTolerance = 8) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (Math.Abs(pixel.R - background.R) > channelTolerance ||
                        Math.Abs(pixel.G - background.G) > channelTolerance ||
                        Math.Abs(pixel.B - background.B) > channelTolerance ||
                        Math.Abs(pixel.A - background.A) > channelTolerance) {
                        count++;
                    }
                }
            }

            return count;
        }

        internal static int CountNonWhiteVisiblePixels(OfficeRasterImage image) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.A == 0) {
                        continue;
                    }

                    if (pixel.R < 245 || pixel.G < 245 || pixel.B < 245 || pixel.A < 250) {
                        count++;
                    }
                }
            }

            return count;
        }

        internal static int ReadNonNegativeInt(string variable, int defaultValue) {
            string? raw = Environment.GetEnvironmentVariable(variable);
            if (string.IsNullOrWhiteSpace(raw)) {
                return defaultValue;
            }

            if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) || value < 0) {
                throw new InvalidOperationException(variable + " must be a non-negative integer.");
            }

            return value;
        }

        internal static double ReadNonNegativeDouble(string variable, double defaultValue) {
            string? raw = Environment.GetEnvironmentVariable(variable);
            if (string.IsNullOrWhiteSpace(raw)) {
                return defaultValue;
            }

            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
                double.IsNaN(value) ||
                value < 0D) {
                throw new InvalidOperationException(variable + " must be a non-negative number.");
            }

            return value;
        }

        internal static string NormalizeText(string text) =>
            text.Replace("\r\n", "\n").Replace('\r', '\n');

        internal static string NormalizeTextWithTrailingNewLine(string text) {
            string normalized = NormalizeText(text);
            return normalized.EndsWith("\n", StringComparison.Ordinal) ? normalized : normalized + "\n";
        }

        internal static string CreateArtifactDirectory(string familyName) {
            string? configuredRoot = Environment.GetEnvironmentVariable("OFFICEIMO_VISUAL_BASELINE_ARTIFACTS_DIRECTORY");
            string root = string.IsNullOrWhiteSpace(configuredRoot)
                ? Path.GetTempPath()
                : configuredRoot;
            string directory = Path.Combine(
                root,
                familyName,
                DateTime.UtcNow.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            return directory;
        }

        internal static string GetTestsProjectRoot() {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null) {
                if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.Excel.Tests.csproj"))) {
                    return directory.FullName;
                }

                if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.Visio.Tests.csproj"))) {
                    return directory.FullName;
                }

                if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.Pdf.Tests.csproj"))) {
                    return directory.FullName;
                }

                string pdfProjectRoot = Path.Combine(directory.FullName, "OfficeIMO.Pdf.Tests");
                if (File.Exists(Path.Combine(pdfProjectRoot, "OfficeIMO.Pdf.Tests.csproj"))) {
                    return pdfProjectRoot;
                }

                directory = directory.Parent;
            }

            throw new DirectoryNotFoundException("Could not locate OfficeIMO test project root from test runtime base directory.");
        }
    }

    internal sealed class VisualRasterComparison {
        internal VisualRasterComparison(
            bool passed,
            int differentPixels,
            int totalPixels,
            int maxChannelDelta,
            int channelTolerance,
            int allowedDifferentPixels,
            double meanAbsoluteError,
            double maximumMeanAbsoluteError,
            double rootMeanSquareError,
            double maximumRootMeanSquareError,
            double meanLuminanceError,
            double maximumMeanLuminanceError,
            byte[] diffPng) {
            Passed = passed;
            DifferentPixels = differentPixels;
            TotalPixels = totalPixels;
            MaxChannelDelta = maxChannelDelta;
            ChannelTolerance = channelTolerance;
            AllowedDifferentPixels = allowedDifferentPixels;
            MeanAbsoluteError = meanAbsoluteError;
            MaximumMeanAbsoluteError = maximumMeanAbsoluteError;
            RootMeanSquareError = rootMeanSquareError;
            MaximumRootMeanSquareError = maximumRootMeanSquareError;
            MeanLuminanceError = meanLuminanceError;
            MaximumMeanLuminanceError = maximumMeanLuminanceError;
            DiffPng = diffPng;
        }

        internal bool Passed { get; }

        internal int DifferentPixels { get; }

        internal int TotalPixels { get; }

        internal int MaxChannelDelta { get; }

        internal int ChannelTolerance { get; }

        internal int AllowedDifferentPixels { get; }

        internal double MeanAbsoluteError { get; }

        internal double MaximumMeanAbsoluteError { get; }

        internal double RootMeanSquareError { get; }

        internal double MaximumRootMeanSquareError { get; }

        internal double MeanLuminanceError { get; }

        internal double MaximumMeanLuminanceError { get; }

        internal byte[] DiffPng { get; }
    }
}
