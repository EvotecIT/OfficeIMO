using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentRasterVisualBaselineTests {
    private static PdfDocument CreateVisualBaselineDocument(PdfOptions options) {
        return PdfDocument.Create(options.UseOfficeFontFamily("Arial"));
    }

    private static void WriteReviewPdfArtifact(string scenarioName, byte[] pdfBytes) {
        string? outputDirectory = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT");
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            return;
        }

        Directory.CreateDirectory(outputDirectory);
        string safeName = scenarioName;
        foreach (char invalidChar in Path.GetInvalidFileNameChars()) {
            safeName = safeName.Replace(invalidChar, '-');
        }

        File.WriteAllBytes(Path.Combine(outputDirectory, safeName + ".pdf"), pdfBytes);
    }

    private static void AssertRasterBaseline(string baselineName, string actualPath) {
        string expectedPath = Path.Combine(GetTestsProjectRoot(), "Pdf", "VisualBaselines", baselineName);
        if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_PDF_RASTER_BASELINE"), "1", StringComparison.Ordinal)) {
            Directory.CreateDirectory(Path.GetDirectoryName(expectedPath)!);
            File.Copy(actualPath, expectedPath, overwrite: true);
            return;
        }

        if (!File.Exists(expectedPath)) {
            throw new FileNotFoundException(
                "PDF raster baseline missing. Set OFFICEIMO_UPDATE_PDF_RASTER_BASELINE=1 and re-run this test to generate it.",
                expectedPath);
        }

        RasterComparison comparison = CompareRasterImages(File.ReadAllBytes(expectedPath), File.ReadAllBytes(actualPath));
        if (!comparison.Passed) {
            string artifactDirectory = Path.Combine(Path.GetTempPath(), "OfficeIMO.PdfRaster", DateTime.UtcNow.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(artifactDirectory);

            string actualArtifactPath = Path.Combine(artifactDirectory, Path.GetFileName(actualPath));
            string expectedArtifactPath = Path.Combine(artifactDirectory, Path.GetFileName(expectedPath));
            string diffArtifactPath = Path.Combine(artifactDirectory, Path.GetFileNameWithoutExtension(actualPath) + ".diff.png");
            File.Copy(actualPath, actualArtifactPath, overwrite: true);
            File.Copy(expectedPath, expectedArtifactPath, overwrite: true);
            File.WriteAllBytes(diffArtifactPath, comparison.DiffPng);

            throw new Xunit.Sdk.XunitException(
                "PDF raster baseline changed. " +
                "Different pixels: " + comparison.DifferentPixels + "/" + comparison.TotalPixels + "; " +
                "max channel delta: " + comparison.MaxChannelDelta + "; " +
                "allowed different pixels: " + comparison.AllowedDifferentPixels + "; " +
                "channel tolerance: " + comparison.ChannelTolerance + ". " +
                "Artifacts: " + artifactDirectory + ".");
        }
    }

    private static RasterComparison CompareRasterImages(byte[] expectedPng, byte[] actualPng) {
        int channelTolerance = ReadNonNegativeInt("OFFICEIMO_PDF_RASTER_PIXEL_TOLERANCE", 0);
        int allowedDifferentPixels = ReadNonNegativeInt("OFFICEIMO_PDF_RASTER_ALLOWED_DIFF_PIXELS", DefaultAllowedRasterNoisePixels);
        return CompareRasterImages(expectedPng, actualPng, channelTolerance, allowedDifferentPixels);
    }

    private static RasterComparison CompareRasterImages(byte[] expectedPng, byte[] actualPng, int channelTolerance, int allowedDifferentPixels) {
        PngRaster expected = PngRaster.Decode(expectedPng);
        PngRaster actual = PngRaster.Decode(actualPng);
        if (expected.Width != actual.Width || expected.Height != actual.Height) {
            byte[] sizeDiff = PngRaster.CreateSizeMismatchDiff(expected, actual);
            return new RasterComparison(false, 0, Math.Max(expected.Width * expected.Height, actual.Width * actual.Height), 255, channelTolerance, allowedDifferentPixels, sizeDiff);
        }

        int differentPixels = 0;
        int maxChannelDelta = 0;
        byte[] diff = new byte[expected.Width * expected.Height * 3];

        for (int pixel = 0; pixel < expected.Width * expected.Height; pixel++) {
            int offset = pixel * 4;
            int deltaR = Math.Abs(expected.Pixels[offset] - actual.Pixels[offset]);
            int deltaG = Math.Abs(expected.Pixels[offset + 1] - actual.Pixels[offset + 1]);
            int deltaB = Math.Abs(expected.Pixels[offset + 2] - actual.Pixels[offset + 2]);
            int deltaA = Math.Abs(expected.Pixels[offset + 3] - actual.Pixels[offset + 3]);
            int maxPixelDelta = Math.Max(Math.Max(deltaR, deltaG), Math.Max(deltaB, deltaA));
            maxChannelDelta = Math.Max(maxChannelDelta, maxPixelDelta);

            int diffOffset = pixel * 3;
            if (maxPixelDelta > channelTolerance) {
                differentPixels++;
                diff[diffOffset] = 255;
                diff[diffOffset + 1] = (byte)Math.Min(255, Math.Max(deltaR, deltaG) * 4);
                diff[diffOffset + 2] = (byte)Math.Min(255, Math.Max(deltaB, deltaA) * 4);
            } else {
                int gray = (expected.Pixels[offset] + expected.Pixels[offset + 1] + expected.Pixels[offset + 2]) / 3;
                byte muted = (byte)(240 - Math.Min(120, gray / 3));
                diff[diffOffset] = muted;
                diff[diffOffset + 1] = muted;
                diff[diffOffset + 2] = muted;
            }
        }

        bool passed = differentPixels <= allowedDifferentPixels;
        return new RasterComparison(passed, differentPixels, expected.Width * expected.Height, maxChannelDelta, channelTolerance, allowedDifferentPixels, PngRaster.EncodeRgb(expected.Width, expected.Height, diff));
    }

    private static int ReadNonNegativeInt(string variable, int defaultValue) {
        string? raw = Environment.GetEnvironmentVariable(variable);
        if (string.IsNullOrWhiteSpace(raw)) {
            return defaultValue;
        }

        int value;
        if (!int.TryParse(raw, out value) || value < 0) {
            throw new InvalidOperationException(variable + " must be a non-negative integer.");
        }

        return value;
    }

    private static void RunPdftoppm(string rasterizerPath, string pdfPath, string outputPrefix, string workDir, int pageNumber) {
        string pageText = pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var psi = new ProcessStartInfo {
            FileName = rasterizerPath,
            Arguments = "-r 72 -png -singlefile -f " + pageText + " -l " + pageText + " " + Quote(pdfPath) + " " + Quote(outputPrefix),
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = workDir
        };

        using var process = Process.Start(psi);
        if (process == null) {
            throw new InvalidOperationException("Could not start PDF rasterizer: " + rasterizerPath);
        }

        if (!process.WaitForExit(30000)) {
            process.Kill();
            throw new TimeoutException("PDF rasterizer timed out: " + rasterizerPath);
        }

        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        if (process.ExitCode != 0) {
            throw new InvalidOperationException("PDF rasterizer failed with exit code " + process.ExitCode + "." + Environment.NewLine + output + Environment.NewLine + error);
        }
    }

    private static bool TryFindPdftoppm(out string path) {
        string? configured = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_RASTERIZER");
        if (!string.IsNullOrWhiteSpace(configured) && File.Exists(configured)) {
            path = configured;
            return true;
        }

        string fileName = Environment.OSVersion.Platform == PlatformID.Win32NT ? "pdftoppm.exe" : "pdftoppm";
        string? pathVariable = Environment.GetEnvironmentVariable("PATH");
        if (!string.IsNullOrWhiteSpace(pathVariable)) {
            foreach (string directory in pathVariable.Split(Path.PathSeparator)) {
                if (string.IsNullOrWhiteSpace(directory)) {
                    continue;
                }

                string candidate = Path.Combine(directory.Trim(), fileName);
                if (File.Exists(candidate)) {
                    path = candidate;
                    return true;
                }
            }
        }

        if (Environment.OSVersion.Platform == PlatformID.Win32NT) {
            string? localAppData = Environment.GetEnvironmentVariable("LOCALAPPDATA");
            if (!string.IsNullOrWhiteSpace(localAppData)) {
                string packages = Path.Combine(localAppData, "Microsoft", "WinGet", "Packages");
                if (Directory.Exists(packages)) {
                    string[] candidates = Directory.GetFiles(packages, "pdftoppm.exe", SearchOption.AllDirectories);
                    if (candidates.Length > 0) {
                        path = candidates[0];
                        return true;
                    }
                }
            }
        }

        path = string.Empty;
        return false;
    }

    private static bool IsRequired() =>
        string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_REQUIRE_PDF_RASTERIZER"), "1", StringComparison.Ordinal);

    private static string Quote(string value) => "\"" + value.Replace("\"", "\\\"") + "\"";

    private sealed class RasterComparison {
        internal RasterComparison(bool passed, int differentPixels, int totalPixels, int maxChannelDelta, int channelTolerance, int allowedDifferentPixels, byte[] diffPng) {
            Passed = passed;
            DifferentPixels = differentPixels;
            TotalPixels = totalPixels;
            MaxChannelDelta = maxChannelDelta;
            ChannelTolerance = channelTolerance;
            AllowedDifferentPixels = allowedDifferentPixels;
            DiffPng = diffPng;
        }

        internal bool Passed { get; }
        internal int DifferentPixels { get; }
        internal int TotalPixels { get; }
        internal int MaxChannelDelta { get; }
        internal int ChannelTolerance { get; }
        internal int AllowedDifferentPixels { get; }
        internal byte[] DiffPng { get; }
    }

    private sealed class PngRaster {
        private static readonly byte[] Signature = { 137, 80, 78, 71, 13, 10, 26, 10 };

        private PngRaster(int width, int height, byte[] pixels) {
            Width = width;
            Height = height;
            Pixels = pixels;
        }

        internal int Width { get; }
        internal int Height { get; }
        internal byte[] Pixels { get; }

        internal static PngRaster Decode(byte[] bytes) {
            if (bytes.Length < Signature.Length || !StartsWith(bytes, Signature)) {
                throw new InvalidOperationException("Raster baseline is not a PNG file.");
            }

            int width = 0;
            int height = 0;
            int bitDepth = 0;
            int colorType = 0;
            int compression = 0;
            int filter = 0;
            int interlace = 0;
            var idat = new List<byte>();

            int offset = Signature.Length;
            while (offset + 12 <= bytes.Length) {
                int length = ReadBigEndianInt32(bytes, offset);
                offset += 4;
                string type = Encoding.ASCII.GetString(bytes, offset, 4);
                offset += 4;
                if (length < 0 || offset + length + 4 > bytes.Length) {
                    throw new InvalidOperationException("PNG chunk length is invalid.");
                }

                if (type == "IHDR") {
                    width = ReadBigEndianInt32(bytes, offset);
                    height = ReadBigEndianInt32(bytes, offset + 4);
                    bitDepth = bytes[offset + 8];
                    colorType = bytes[offset + 9];
                    compression = bytes[offset + 10];
                    filter = bytes[offset + 11];
                    interlace = bytes[offset + 12];
                } else if (type == "IDAT") {
                    for (int i = 0; i < length; i++) {
                        idat.Add(bytes[offset + i]);
                    }
                } else if (type == "IEND") {
                    break;
                }

                offset += length + 4;
            }

            if (width <= 0 || height <= 0) {
                throw new InvalidOperationException("PNG image dimensions are invalid.");
            }

            if (bitDepth != 8 || compression != 0 || filter != 0 || interlace != 0 || (colorType != 2 && colorType != 6)) {
                throw new InvalidOperationException("Only non-interlaced 8-bit RGB/RGBA PNG raster baselines are supported.");
            }

            byte[] inflated = InflateZlib(idat.ToArray());
            int channels = colorType == 6 ? 4 : 3;
            int stride = width * channels;
            byte[] pixels = new byte[width * height * 4];
            byte[] previous = new byte[stride];
            byte[] current = new byte[stride];
            int source = 0;
            int bytesPerPixel = channels;

            for (int y = 0; y < height; y++) {
                if (source >= inflated.Length) {
                    throw new InvalidOperationException("PNG image data ended unexpectedly.");
                }

                byte filterType = inflated[source++];
                if (source + stride > inflated.Length) {
                    throw new InvalidOperationException("PNG scanline is incomplete.");
                }

                Buffer.BlockCopy(inflated, source, current, 0, stride);
                source += stride;
                UnfilterScanline(filterType, current, previous, bytesPerPixel);

                for (int x = 0; x < width; x++) {
                    int sourcePixel = x * channels;
                    int targetPixel = (y * width + x) * 4;
                    pixels[targetPixel] = current[sourcePixel];
                    pixels[targetPixel + 1] = current[sourcePixel + 1];
                    pixels[targetPixel + 2] = current[sourcePixel + 2];
                    pixels[targetPixel + 3] = colorType == 6 ? current[sourcePixel + 3] : (byte)255;
                }

                byte[] swap = previous;
                previous = current;
                current = swap;
                Array.Clear(current, 0, current.Length);
            }

            return new PngRaster(width, height, pixels);
        }

        internal static byte[] EncodeRgb(int width, int height, byte[] rgb) {
            if (width <= 0 || height <= 0) {
                throw new ArgumentOutOfRangeException(nameof(width), "PNG dimensions must be positive.");
            }

            if (rgb.Length != width * height * 3) {
                throw new ArgumentException("RGB buffer length does not match PNG dimensions.", nameof(rgb));
            }

            byte[] scanlines = new byte[height * (1 + width * 3)];
            int source = 0;
            int target = 0;
            for (int y = 0; y < height; y++) {
                scanlines[target++] = 0;
                Buffer.BlockCopy(rgb, source, scanlines, target, width * 3);
                source += width * 3;
                target += width * 3;
            }

            using var ms = new MemoryStream();
            ms.Write(Signature, 0, Signature.Length);
            byte[] ihdr = new byte[13];
            WriteBigEndianInt32(ihdr, 0, width);
            WriteBigEndianInt32(ihdr, 4, height);
            ihdr[8] = 8;
            ihdr[9] = 2;
            WriteChunk(ms, "IHDR", ihdr);
            WriteChunk(ms, "IDAT", DeflateZlibStored(scanlines));
            WriteChunk(ms, "IEND", Array.Empty<byte>());
            return ms.ToArray();
        }

        internal static byte[] CreateSizeMismatchDiff(PngRaster expected, PngRaster actual) {
            int width = Math.Max(expected.Width, actual.Width);
            int height = Math.Max(expected.Height, actual.Height);
            byte[] diff = new byte[width * height * 3];
            for (int i = 0; i < diff.Length; i += 3) {
                diff[i] = 255;
                diff[i + 1] = 0;
                diff[i + 2] = 255;
            }

            return EncodeRgb(width, height, diff);
        }

        private static void UnfilterScanline(byte filterType, byte[] current, byte[] previous, int bytesPerPixel) {
            for (int i = 0; i < current.Length; i++) {
                int left = i >= bytesPerPixel ? current[i - bytesPerPixel] : 0;
                int up = previous[i];
                int upLeft = i >= bytesPerPixel ? previous[i - bytesPerPixel] : 0;
                int predictor;
                switch (filterType) {
                    case 0:
                        predictor = 0;
                        break;
                    case 1:
                        predictor = left;
                        break;
                    case 2:
                        predictor = up;
                        break;
                    case 3:
                        predictor = (left + up) / 2;
                        break;
                    case 4:
                        predictor = Paeth(left, up, upLeft);
                        break;
                    default:
                        throw new InvalidOperationException("Unsupported PNG scanline filter: " + filterType + ".");
                }

                current[i] = (byte)((current[i] + predictor) & 0xFF);
            }
        }

        private static int Paeth(int left, int up, int upLeft) {
            int p = left + up - upLeft;
            int pa = Math.Abs(p - left);
            int pb = Math.Abs(p - up);
            int pc = Math.Abs(p - upLeft);
            if (pa <= pb && pa <= pc) return left;
            return pb <= pc ? up : upLeft;
        }

        private static bool StartsWith(byte[] bytes, byte[] prefix) {
            for (int i = 0; i < prefix.Length; i++) {
                if (bytes[i] != prefix[i]) {
                    return false;
                }
            }

            return true;
        }

        private static byte[] InflateZlib(byte[] zlib) {
            if (zlib.Length < 6) {
                throw new InvalidOperationException("PNG zlib stream is too short.");
            }

            using var source = new MemoryStream(zlib, 2, zlib.Length - 6);
            using var deflate = new DeflateStream(source, CompressionMode.Decompress);
            using var output = new MemoryStream();
            deflate.CopyTo(output);
            return output.ToArray();
        }

        private static byte[] DeflateZlibStored(byte[] data) {
            using var ms = new MemoryStream();
            ms.WriteByte(0x78);
            ms.WriteByte(0x01);

            int offset = 0;
            while (offset < data.Length) {
                int blockLength = Math.Min(65535, data.Length - offset);
                bool final = offset + blockLength >= data.Length;
                ms.WriteByte(final ? (byte)1 : (byte)0);
                ms.WriteByte((byte)(blockLength & 0xFF));
                ms.WriteByte((byte)((blockLength >> 8) & 0xFF));
                int nlen = blockLength ^ 0xFFFF;
                ms.WriteByte((byte)(nlen & 0xFF));
                ms.WriteByte((byte)((nlen >> 8) & 0xFF));
                ms.Write(data, offset, blockLength);
                offset += blockLength;
            }

            uint adler = Adler32(data);
            ms.WriteByte((byte)((adler >> 24) & 0xFF));
            ms.WriteByte((byte)((adler >> 16) & 0xFF));
            ms.WriteByte((byte)((adler >> 8) & 0xFF));
            ms.WriteByte((byte)(adler & 0xFF));
            return ms.ToArray();
        }

        private static uint Adler32(byte[] data) {
            const uint mod = 65521;
            uint a = 1;
            uint b = 0;
            for (int i = 0; i < data.Length; i++) {
                a = (a + data[i]) % mod;
                b = (b + a) % mod;
            }

            return (b << 16) | a;
        }

        private static int ReadBigEndianInt32(byte[] bytes, int offset) =>
            (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

        private static void WriteBigEndianInt32(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)((value >> 24) & 0xFF);
            bytes[offset + 1] = (byte)((value >> 16) & 0xFF);
            bytes[offset + 2] = (byte)((value >> 8) & 0xFF);
            bytes[offset + 3] = (byte)(value & 0xFF);
        }

        private static void WriteChunk(Stream stream, string type, byte[] data) {
            byte[] typeBytes = Encoding.ASCII.GetBytes(type);
            byte[] length = new byte[4];
            WriteBigEndianInt32(length, 0, data.Length);
            stream.Write(length, 0, length.Length);
            stream.Write(typeBytes, 0, typeBytes.Length);
            stream.Write(data, 0, data.Length);

            uint crc = Crc32(typeBytes, data);
            byte[] crcBytes = new byte[4];
            WriteBigEndianInt32(crcBytes, 0, unchecked((int)crc));
            stream.Write(crcBytes, 0, crcBytes.Length);
        }

        private static uint Crc32(byte[] type, byte[] data) {
            uint crc = 0xFFFFFFFF;
            for (int i = 0; i < type.Length; i++) {
                crc = UpdateCrc(crc, type[i]);
            }

            for (int i = 0; i < data.Length; i++) {
                crc = UpdateCrc(crc, data[i]);
            }

            return crc ^ 0xFFFFFFFF;
        }

        private static uint UpdateCrc(uint crc, byte value) {
            crc ^= value;
            for (int i = 0; i < 8; i++) {
                crc = (crc & 1) != 0 ? 0xEDB88320 ^ (crc >> 1) : crc >> 1;
            }

            return crc;
        }
    }

    private static string GetTestsProjectRoot() {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.Tests.csproj"))) {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate OfficeIMO.Tests project root from test runtime base directory.");
    }

    private static void TryDeleteDirectory(string directory) {
        try {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        } catch {
        }
    }
}
