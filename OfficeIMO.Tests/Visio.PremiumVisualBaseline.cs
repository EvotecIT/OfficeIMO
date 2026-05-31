using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioPremiumVisualBaselineTests {
        private static readonly IReadOnlyDictionary<string, string> BaselinePrefixes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["Premium Cloud Architecture"] = "officeimo-visio-premium-cloud-architecture",
            ["Premium Network Segmentation"] = "officeimo-visio-premium-network-segmentation",
            ["Premium Executive Dependencies"] = "officeimo-visio-premium-executive-dependencies",
            ["Premium Technical Topology"] = "officeimo-visio-premium-technical-topology",
            ["Premium Print Audit Trail"] = "officeimo-visio-premium-print-audit-trail",
            ["Premium Incident Sequence"] = "officeimo-visio-premium-incident-sequence",
            ["Premium Release Timeline"] = "officeimo-visio-premium-release-timeline",
            ["Premium Governed Process"] = "officeimo-visio-premium-governed-process"
        };

        [Fact]
        public void PremiumGalleryPreviewsMatchApprovedBaselines() {
            if (!VisioDesktopValidator.IsAvailable()) {
                if (IsRequired()) {
                    throw new InvalidOperationException("Premium Visio visual baseline tests require Microsoft Visio desktop automation. Install Visio or unset OFFICEIMO_REQUIRE_VISIO_PREMIUM_BASELINES.");
                }

                return;
            }

            string workDirectory = Path.Combine(Path.GetTempPath(), "OfficeIMO.VisioPremiumBaselines", Guid.NewGuid().ToString("N"));
            string documentsDirectory = Path.Combine(workDirectory, "documents");
            string actualDirectory = Path.Combine(workDirectory, "actual");

            try {
                IReadOnlyList<VisioGalleryResult> results = VisioPremiumGallery.Create(documentsDirectory);

                Assert.Equal(BaselinePrefixes.Count, results.Count);
                Assert.All(results, result => Assert.True(result.IsClean, FormatGalleryIssues(result)));

                foreach (VisioGalleryResult result in results.OrderBy(result => result.Name, StringComparer.Ordinal)) {
                    string prefix = GetBaselinePrefix(result.Name);
                    VisioPremiumBaselineContext context = CreateBaselineContext(result, prefix);
                    AssertTextBaseline(context.InspectionBaselineName, context.InspectionText);
                    AssertTextBaseline(context.StencilProfileBaselineName, context.StencilProfileText);

                    VisioDesktopValidationOptions options = new() {
                        ExportDirectory = actualDirectory,
                        ExportFileNamePrefix = prefix
                    };
                    options.ExportFormats.Add(VisioDesktopExportFormat.Png);
                    options.ExportFormats.Add(VisioDesktopExportFormat.Svg);

                    VisioDesktopValidationResult desktop = VisioDesktopValidator.Validate(result.FilePath, options);
                    Assert.True(desktop.IsValid, string.Join(Environment.NewLine, desktop.Issues));
                    Assert.Equal(2, desktop.OutputFiles.Count);

                    foreach (string actualPath in desktop.OutputFiles.OrderBy(path => path, StringComparer.OrdinalIgnoreCase)) {
                        AssertBaseline(Path.GetFileName(actualPath), actualPath, context);
                    }
                }
            } finally {
                TryDeleteDirectory(workDirectory);
            }
        }

        [Fact]
        public void PngBaselineComparisonReportsPixelDiffAndProducesDiffPng() {
            byte[] expected = PngRaster.EncodeRgb(2, 1, new byte[] {
                255, 255, 255,
                0, 0, 0
            });
            byte[] actual = PngRaster.EncodeRgb(2, 1, new byte[] {
                255, 255, 255,
                0, 64, 255
            });

            RasterComparison comparison = CompareRasterImages(expected, actual, channelTolerance: 0, allowedDifferentPixels: 0);

            Assert.False(comparison.Passed);
            Assert.Equal(1, comparison.DifferentPixels);
            Assert.Equal(2, comparison.TotalPixels);
            Assert.Equal(255, comparison.MaxChannelDelta);
            Assert.True(comparison.DiffPng.Length > 0);
            Assert.Equal(2, PngRaster.Decode(comparison.DiffPng).Width);
        }

        [Fact]
        public void TextDiffReportsAddedRemovedAndChangedLines() {
            string expected = string.Join(Environment.NewLine, new[] {
                "document.title=Before",
                "page[Main].shape[1].text=Keep",
                "page[Main].shape[2].text=Remove"
            });
            string actual = string.Join(Environment.NewLine, new[] {
                "document.title=After",
                "page[Main].shape[1].text=Keep",
                "page[Main].shape[3].text=Add"
            });

            string diff = CreateLineDiff(expected, actual);

            Assert.Contains("Changed document.title expected=Before actual=After", diff, StringComparison.Ordinal);
            Assert.Contains("Removed page[Main].shape[2].text expected=Remove actual=", diff, StringComparison.Ordinal);
            Assert.Contains("Added page[Main].shape[3].text expected= actual=Add", diff, StringComparison.Ordinal);
        }

        private static string GetBaselinePrefix(string name) {
            if (BaselinePrefixes.TryGetValue(name, out string? prefix)) {
                return prefix;
            }

            throw new InvalidOperationException("No premium Visio visual baseline prefix is registered for '" + name + "'.");
        }

        private static VisioPremiumBaselineContext CreateBaselineContext(VisioGalleryResult result, string prefix) {
            VisioDocument loaded = VisioDocument.Load(result.FilePath);
            VisioInspectionSnapshot inspection = loaded.CreateInspectionSnapshot();
            VisioStencilProfile stencilProfile = inspection.CreateStencilProfile();
            return new VisioPremiumBaselineContext(
                prefix + ".inspection.txt",
                inspection.ToText(),
                prefix + ".stencil-profile.txt",
                stencilProfile.ToText());
        }

        private static void AssertBaseline(string baselineName, string actualPath, VisioPremiumBaselineContext context) {
            string expectedPath = Path.Combine(GetTestsProjectRoot(), "Visio", "VisualBaselines", baselineName);
            if (IsBaselineUpdateRequested()) {
                Directory.CreateDirectory(Path.GetDirectoryName(expectedPath)!);
                File.Copy(actualPath, expectedPath, overwrite: true);
                return;
            }

            if (!File.Exists(expectedPath)) {
                throw new FileNotFoundException(
                    "Premium Visio visual baseline missing. Set OFFICEIMO_UPDATE_VISIO_PREMIUM_BASELINES=1 and re-run this test to generate it.",
                    expectedPath);
            }

            if (string.Equals(Path.GetExtension(expectedPath), ".svg", StringComparison.OrdinalIgnoreCase)) {
                if (string.Equals(CanonicalizeSvg(expectedPath), CanonicalizeSvg(actualPath), StringComparison.Ordinal)) {
                    return;
                }

                ThrowBaselineChanged(baselineName, expectedPath, actualPath, null, context);
                return;
            }

            if (!string.Equals(Path.GetExtension(expectedPath), ".png", StringComparison.OrdinalIgnoreCase)) {
                if (File.ReadAllBytes(expectedPath).AsSpan().SequenceEqual(File.ReadAllBytes(actualPath))) {
                    return;
                }

                ThrowBaselineChanged(baselineName, expectedPath, actualPath, null, context);
                return;
            }

            RasterComparison comparison = CompareRasterImages(File.ReadAllBytes(expectedPath), File.ReadAllBytes(actualPath));
            if (comparison.Passed) {
                return;
            }

            ThrowBaselineChanged(baselineName, expectedPath, actualPath, comparison, context);
        }

        private static void AssertTextBaseline(string baselineName, string actualText) {
            string expectedPath = Path.Combine(GetTestsProjectRoot(), "Visio", "VisualBaselines", baselineName);
            if (IsStructuralBaselineUpdateRequested()) {
                Directory.CreateDirectory(Path.GetDirectoryName(expectedPath)!);
                File.WriteAllText(expectedPath, NormalizeText(actualText), new UTF8Encoding(false));
                return;
            }

            if (!File.Exists(expectedPath)) {
                throw new FileNotFoundException(
                    "Premium Visio structural baseline missing. Set OFFICEIMO_UPDATE_VISIO_PREMIUM_STRUCTURAL_BASELINES=1 or OFFICEIMO_UPDATE_VISIO_PREMIUM_BASELINES=1 and re-run this test to generate it.",
                    expectedPath);
            }

            string expectedText = NormalizeText(File.ReadAllText(expectedPath));
            string normalizedActual = NormalizeText(actualText);
            if (string.Equals(expectedText, normalizedActual, StringComparison.Ordinal)) {
                return;
            }

            ThrowTextBaselineChanged(baselineName, expectedPath, expectedText, normalizedActual);
        }

        private static void ThrowBaselineChanged(string baselineName, string expectedPath, string actualPath, RasterComparison? comparison, VisioPremiumBaselineContext context) {
            string artifactDirectory = CreateArtifactDirectory();

            File.Copy(expectedPath, Path.Combine(artifactDirectory, "expected-" + Path.GetFileName(expectedPath)), overwrite: true);
            File.Copy(actualPath, Path.Combine(artifactDirectory, "actual-" + Path.GetFileName(actualPath)), overwrite: true);
            if (comparison != null) {
                File.WriteAllBytes(
                    Path.Combine(artifactDirectory, Path.GetFileNameWithoutExtension(actualPath) + ".diff.png"),
                    comparison.DiffPng);
            }

            WriteContextArtifacts(artifactDirectory, context);

            FileInfo expectedInfo = new(expectedPath);
            FileInfo actualInfo = new(actualPath);
            string pixelSummary = comparison == null
                ? string.Empty
                : "Different pixels: " + comparison.DifferentPixels + "/" + comparison.TotalPixels + "; " +
                  "max channel delta: " + comparison.MaxChannelDelta + "; " +
                  "allowed different pixels: " + comparison.AllowedDifferentPixels + "; " +
                  "channel tolerance: " + comparison.ChannelTolerance + ". ";
            throw new Xunit.Sdk.XunitException(
                "Premium Visio visual baseline changed for '" + baselineName + "'. " +
                "Expected bytes: " + expectedInfo.Length + "; actual bytes: " + actualInfo.Length + ". " +
                pixelSummary +
                "Artifacts: " + artifactDirectory + ".");
        }

        private static void ThrowTextBaselineChanged(string baselineName, string expectedPath, string expectedText, string actualText) {
            string artifactDirectory = CreateArtifactDirectory();
            File.Copy(expectedPath, Path.Combine(artifactDirectory, "expected-" + Path.GetFileName(expectedPath)), overwrite: true);
            File.WriteAllText(Path.Combine(artifactDirectory, "actual-" + baselineName), actualText, new UTF8Encoding(false));
            File.WriteAllText(Path.Combine(artifactDirectory, Path.GetFileNameWithoutExtension(baselineName) + ".diff.txt"), CreateLineDiff(expectedText, actualText), new UTF8Encoding(false));

            throw new Xunit.Sdk.XunitException(
                "Premium Visio structural baseline changed for '" + baselineName + "'. " +
                "Artifacts: " + artifactDirectory + ".");
        }

        private static void WriteContextArtifacts(string artifactDirectory, VisioPremiumBaselineContext context) {
            WriteContextArtifact(artifactDirectory, context.InspectionBaselineName, context.InspectionText);
            WriteContextArtifact(artifactDirectory, context.StencilProfileBaselineName, context.StencilProfileText);
        }

        private static void WriteContextArtifact(string artifactDirectory, string baselineName, string actualText) {
            string expectedPath = Path.Combine(GetTestsProjectRoot(), "Visio", "VisualBaselines", baselineName);
            string normalizedActual = NormalizeText(actualText);
            File.WriteAllText(Path.Combine(artifactDirectory, "actual-" + baselineName), normalizedActual, new UTF8Encoding(false));
            if (!File.Exists(expectedPath)) {
                return;
            }

            string expectedText = NormalizeText(File.ReadAllText(expectedPath));
            File.Copy(expectedPath, Path.Combine(artifactDirectory, "expected-" + baselineName), overwrite: true);
            File.WriteAllText(Path.Combine(artifactDirectory, Path.GetFileNameWithoutExtension(baselineName) + ".diff.txt"), CreateLineDiff(expectedText, normalizedActual), new UTF8Encoding(false));
        }

        private static string CreateArtifactDirectory() {
            string artifactDirectory = Path.Combine(
                Path.GetTempPath(),
                "OfficeIMO.VisioPremiumBaselines",
                DateTime.UtcNow.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(artifactDirectory);
            return artifactDirectory;
        }

        private static string CreateLineDiff(string expectedText, string actualText) {
            SortedDictionary<string, string> expectedLines = ToLineMap(expectedText);
            SortedDictionary<string, string> actualLines = ToLineMap(actualText);
            SortedSet<string> keys = new(expectedLines.Keys, StringComparer.Ordinal);
            keys.UnionWith(actualLines.Keys);

            StringBuilder builder = new();
            foreach (string key in keys) {
                bool hasExpected = expectedLines.TryGetValue(key, out string? expectedValue);
                bool hasActual = actualLines.TryGetValue(key, out string? actualValue);
                if (!hasExpected && hasActual) {
                    builder.Append("Added ");
                } else if (hasExpected && !hasActual) {
                    builder.Append("Removed ");
                } else if (!string.Equals(expectedValue, actualValue, StringComparison.Ordinal)) {
                    builder.Append("Changed ");
                } else {
                    continue;
                }

                builder.Append(key);
                builder.Append(" expected=");
                builder.Append(expectedValue);
                builder.Append(" actual=");
                builder.Append(actualValue);
                builder.AppendLine();
            }

            return builder.ToString();
        }

        private static SortedDictionary<string, string> ToLineMap(string text) {
            SortedDictionary<string, string> map = new(StringComparer.Ordinal);
            string[] lines = NormalizeText(text).Split('\n');
            foreach (string line in lines) {
                if (line.Length == 0) {
                    continue;
                }

                int separator = line.IndexOf('=');
                string key = separator >= 0 ? line.Substring(0, separator) : line;
                string value = separator >= 0 ? line.Substring(separator + 1) : string.Empty;
                map[key] = value;
            }

            return map;
        }

        private static string NormalizeText(string text) {
            string normalized = text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace("\r", "\n", StringComparison.Ordinal);
            return normalized.EndsWith("\n", StringComparison.Ordinal) ? normalized : normalized + "\n";
        }

        private static RasterComparison CompareRasterImages(byte[] expectedPng, byte[] actualPng) {
            int channelTolerance = ReadNonNegativeInt("OFFICEIMO_VISIO_PREMIUM_BASELINE_PIXEL_TOLERANCE", 0);
            int allowedDifferentPixels = ReadNonNegativeInt("OFFICEIMO_VISIO_PREMIUM_BASELINE_ALLOWED_DIFF_PIXELS", 0);
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

        private static string CanonicalizeSvg(string path) {
            string svg = File.ReadAllText(path)
                .Replace("\r\n", "\n", StringComparison.Ordinal)
                .Replace("\r", "\n", StringComparison.Ordinal);

            Dictionary<string, string> styleByClass = new(StringComparer.Ordinal);
            SortedSet<string> uniqueStyles = new(StringComparer.Ordinal);
            foreach (Match match in Regex.Matches(svg, @"\.st(?<id>\d+)\s*\{(?<style>[^}]*)\}", RegexOptions.CultureInvariant)) {
                string className = "st" + match.Groups["id"].Value;
                string style = NormalizeSvgStyle(match.Groups["style"].Value);
                styleByClass[className] = style;
                uniqueStyles.Add(style);
            }

            Dictionary<string, string> canonicalClassByStyle = new(StringComparer.Ordinal);
            int index = 1;
            foreach (string style in uniqueStyles) {
                canonicalClassByStyle[style] = "style" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
                index++;
            }

            Dictionary<string, string> canonicalClassByOriginalClass = styleByClass.ToDictionary(
                pair => pair.Key,
                pair => canonicalClassByStyle[pair.Value],
                StringComparer.Ordinal);

            svg = Regex.Replace(svg, @"<style\b[^>]*>.*?</style>\s*", string.Empty, RegexOptions.Singleline | RegexOptions.CultureInvariant);
            svg = Regex.Replace(
                svg,
                @"\bst(?<id>\d+)\b",
                match => canonicalClassByOriginalClass.TryGetValue("st" + match.Groups["id"].Value, out string? canonicalClass)
                    ? canonicalClass
                    : match.Value,
                RegexOptions.CultureInvariant);

            return svg.Trim();
        }

        private static string NormalizeSvgStyle(string style) {
            string normalized = Regex.Replace(style.Trim(), @"\s+", " ", RegexOptions.CultureInvariant);
            if (normalized.Contains("stroke:none", StringComparison.Ordinal)) {
                normalized = Regex.Replace(normalized, @";?stroke-width:[^;]+", string.Empty, RegexOptions.CultureInvariant);
            }

            return normalized.Trim(';');
        }

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

        private sealed class VisioPremiumBaselineContext {
            internal VisioPremiumBaselineContext(string inspectionBaselineName, string inspectionText, string stencilProfileBaselineName, string stencilProfileText) {
                InspectionBaselineName = inspectionBaselineName;
                InspectionText = inspectionText;
                StencilProfileBaselineName = stencilProfileBaselineName;
                StencilProfileText = stencilProfileText;
            }

            internal string InspectionBaselineName { get; }
            internal string InspectionText { get; }
            internal string StencilProfileBaselineName { get; }
            internal string StencilProfileText { get; }
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
                    throw new InvalidOperationException("Visio premium baseline is not a PNG file.");
                }

                int width = 0;
                int height = 0;
                int bitDepth = 0;
                int colorType = 0;
                int compression = 0;
                int filter = 0;
                int interlace = 0;
                List<byte> idat = new();

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
                    throw new InvalidOperationException("Only non-interlaced 8-bit RGB/RGBA PNG premium baselines are supported.");
                }

                byte[] inflated = InflateZlib(idat.ToArray());
                int channels = colorType == 6 ? 4 : 3;
                int stride = width * channels;
                byte[] pixels = new byte[width * height * 4];
                byte[] previous = new byte[stride];
                byte[] current = new byte[stride];
                int source = 0;

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
                    UnfilterScanline(filterType, current, previous, channels);

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

                using MemoryStream ms = new();
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
                if (pa <= pb && pa <= pc) {
                    return left;
                }

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

                using MemoryStream source = new(zlib, 2, zlib.Length - 6);
                using DeflateStream deflate = new(source, CompressionMode.Decompress);
                using MemoryStream output = new();
                deflate.CopyTo(output);
                return output.ToArray();
            }

            private static byte[] DeflateZlibStored(byte[] data) {
                using MemoryStream ms = new();
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

        private static string FormatGalleryIssues(VisioGalleryResult result) {
            IEnumerable<string> issues = result.PackageIssues
                .Concat(result.QualityIssues.Select(issue => issue.ToString()))
                .Concat(result.DesktopValidation?.Issues ?? Array.Empty<string>());
            return result.Name + Environment.NewLine + string.Join(Environment.NewLine, issues);
        }

        private static bool IsRequired() =>
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_REQUIRE_VISIO_PREMIUM_BASELINES"), "1", StringComparison.Ordinal);

        private static bool IsBaselineUpdateRequested() =>
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_VISIO_PREMIUM_BASELINES"), "1", StringComparison.Ordinal);

        private static bool IsStructuralBaselineUpdateRequested() =>
            IsBaselineUpdateRequested() ||
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_VISIO_PREMIUM_STRUCTURAL_BASELINES"), "1", StringComparison.Ordinal);

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
}
