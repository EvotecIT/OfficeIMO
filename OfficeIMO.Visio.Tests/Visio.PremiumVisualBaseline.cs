using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioPremiumVisualBaselineTests {
        private const string PrintAuditTrailNativePngBaseline = "officeimo-visio-premium-print-audit-trail-native-page1.png";
        private const string LinuxPrintAuditTrailNativePngBaseline = "officeimo-visio-premium-print-audit-trail-native-page1.linux.png";

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

        public static IEnumerable<object[]> PremiumGalleryScenarios {
            get {
                foreach (string name in VisioPremiumGallery.ScenarioNames) {
                    yield return new object[] { name };
                }
            }
        }

        [Theory]
        [MemberData(nameof(PremiumGalleryScenarios))]
        public void PremiumGalleryPreviewMatchesApprovedBaselines(string scenarioName) {
            if (!IsDesktopBaselineRunRequested()) {
                return;
            }

            if (!VisioDesktopBaselineValidator.IsAvailable()) {
                if (IsRequired() || IsBaselineUpdateRequested()) {
                    throw new InvalidOperationException("Premium Visio visual baseline tests require Microsoft Visio desktop automation. Install Visio or unset OFFICEIMO_REQUIRE_VISIO_PREMIUM_BASELINES / OFFICEIMO_UPDATE_VISIO_PREMIUM_BASELINES.");
                }

                return;
            }

            string workDirectory = Path.Combine(Path.GetTempPath(), "OfficeIMO.VisioPremiumBaselines", Guid.NewGuid().ToString("N"));
            string documentsDirectory = Path.Combine(workDirectory, "documents");
            string actualDirectory = Path.Combine(workDirectory, "actual");

            try {
                VisioGalleryResult result = VisioPremiumGallery.CreateScenario(documentsDirectory, scenarioName);
                Assert.True(result.IsClean, FormatGalleryIssues(result));

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

                VisioDesktopValidationResult desktop = VisioDesktopBaselineValidator.Validate(result.FilePath, options);
                Assert.True(desktop.IsValid, string.Join(Environment.NewLine, desktop.Issues));
                Assert.Equal(2, desktop.OutputFiles.Count);

                foreach (string actualPath in desktop.OutputFiles.OrderBy(path => path, StringComparer.OrdinalIgnoreCase)) {
                    AssertBaseline(Path.GetFileName(actualPath), actualPath, context);
                }
            } finally {
                TryDeleteDirectory(workDirectory);
            }
        }

        [Theory]
        [MemberData(nameof(PremiumGalleryScenarios))]
        public void PremiumGalleryNativePreviewMatchesApprovedBaselines(string scenarioName) {
            string workDirectory = Path.Combine(Path.GetTempPath(), "OfficeIMO.VisioPremiumNativeBaselines", Guid.NewGuid().ToString("N"));
            string documentsDirectory = Path.Combine(workDirectory, "documents");
            string actualDirectory = Path.Combine(workDirectory, "actual");

            try {
                VisioGalleryResult result = VisioPremiumGallery.CreateScenario(documentsDirectory, scenarioName);
                Assert.True(result.IsClean, FormatGalleryIssues(result));
                Directory.CreateDirectory(actualDirectory);
                string prefix = GetBaselinePrefix(result.Name);
                VisioPremiumBaselineContext context = CreateBaselineContext(result, prefix);
                VisioDocument document = VisioDocument.Load(result.FilePath);

                string svgPath = Path.Combine(actualDirectory, prefix + "-native-page1.svg");
                document.SaveAsSvg(svgPath, new VisioSvgSaveOptions {
                    PageIndex = 0,
                    PixelsPerInch = 96
                });
                AssertBaseline(Path.GetFileName(svgPath), svgPath, context);

                string pngPath = Path.Combine(actualDirectory, prefix + "-native-page1.png");
                document.SaveAsPng(pngPath, new VisioPngSaveOptions {
                    PageIndex = 0,
                    PixelsPerInch = 96,
                    Supersampling = 3
                });
                AssertBaseline(Path.GetFileName(pngPath), pngPath, context);
            } finally {
                TryDeleteDirectory(workDirectory);
            }
        }

        [Fact]
        public void PremiumGalleryNativeApprovedBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = Path.Combine(VisualBaselineTestSupport.GetTestsProjectRoot(), "Visio", "VisualBaselines");
            foreach (string prefix in BaselinePrefixes.Values.OrderBy(prefix => prefix, StringComparer.Ordinal)) {
                string svgPath = Path.Combine(baselineDirectory, prefix + "-native-page1.svg");
                string pngPath = Path.Combine(baselineDirectory, prefix + "-native-page1.png");
                Assert.True(File.Exists(svgPath), "Missing approved native SVG baseline: " + svgPath);
                Assert.True(File.Exists(pngPath), "Missing approved native PNG baseline: " + pngPath);

                AssertNativeSvgBaselineIsRenderable(svgPath);
                AssertNativePngBaselineIsNonBlank(pngPath);
            }

            string linuxPngPath = Path.Combine(baselineDirectory, LinuxPrintAuditTrailNativePngBaseline);
            Assert.True(File.Exists(linuxPngPath), "Missing approved Linux native PNG baseline: " + linuxPngPath);
            AssertNativePngBaselineIsNonBlank(linuxPngPath);
        }

        [Fact]
        public void PngBaselineComparisonReportsPixelDiffAndProducesDiffPng() {
            byte[] expected = VisualBaselineTestSupport.CreateRgbPng(2, 1, new byte[] {
                255, 255, 255,
                0, 0, 0
            });
            byte[] actual = VisualBaselineTestSupport.CreateRgbPng(2, 1, new byte[] {
                255, 255, 255,
                0, 64, 255
            });

            VisualRasterComparison comparison = CompareRasterImages(expected, actual, channelTolerance: 0, allowedDifferentPixels: 0);

            Assert.False(comparison.Passed);
            Assert.Equal(1, comparison.DifferentPixels);
            Assert.Equal(2, comparison.TotalPixels);
            Assert.Equal(255, comparison.MaxChannelDelta);
            Assert.Equal(39.875D, comparison.MeanAbsoluteError, 3);
            Assert.InRange(comparison.RootMeanSquareError, 92.9D, 93D);
            Assert.InRange(comparison.MeanLuminanceError, 32D, 32.2D);
            Assert.True(comparison.DiffPng.Length > 0);
            Assert.Equal(2, VisualBaselineTestSupport.DecodePng(comparison.DiffPng, "Visio diff PNG is not a supported PNG file.").Width);
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
            string expectedBaselineName = ResolveExpectedBaselineName(baselineName);
            string expectedPath = Path.Combine(VisualBaselineTestSupport.GetTestsProjectRoot(), "Visio", "VisualBaselines", expectedBaselineName);
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

            bool allowNativeVariance =
                string.Equals(expectedBaselineName, baselineName, StringComparison.OrdinalIgnoreCase) &&
                IsNativePngBaseline(baselineName);
            VisualRasterComparison comparison = CompareRasterImages(File.ReadAllBytes(expectedPath), File.ReadAllBytes(actualPath), allowNativeVariance);
            if (comparison.Passed) {
                return;
            }

            ThrowBaselineChanged(baselineName, expectedPath, actualPath, comparison, context);
        }

        private static void AssertTextBaseline(string baselineName, string actualText) {
            string expectedPath = Path.Combine(VisualBaselineTestSupport.GetTestsProjectRoot(), "Visio", "VisualBaselines", baselineName);
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

        private static void AssertNativeSvgBaselineIsRenderable(string svgPath) {
            XDocument document = XDocument.Load(svgPath);
            XNamespace ns = "http://www.w3.org/2000/svg";
            Assert.Equal(ns + "svg", document.Root!.Name);

            int visibleElementCount = document.Root
                .Descendants()
                .Count(element =>
                    element.Name == ns + "path" ||
                    element.Name == ns + "rect" ||
                    element.Name == ns + "ellipse" ||
                    element.Name == ns + "line" ||
                    element.Name == ns + "polyline" ||
                    element.Name == ns + "polygon" ||
                    element.Name == ns + "text");

            Assert.True(visibleElementCount >= 5, "Native SVG baseline appears too sparse to be a rendered gallery preview: " + svgPath);
        }

        private static void AssertNativePngBaselineIsNonBlank(string pngPath) {
            OfficeRasterImage raster = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved native Visio PNG baseline is not a supported PNG file.");
            Assert.True(raster.Width >= 200, "Native PNG baseline width is unexpectedly small: " + pngPath);
            Assert.True(raster.Height >= 150, "Native PNG baseline height is unexpectedly small: " + pngPath);

            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonWhiteVisiblePixels(raster);
            int totalPixels = raster.Width * raster.Height;
            int minimumVisiblePixels = Math.Max(250, totalPixels / 200);
            Assert.True(
                nonBackgroundPixels >= minimumVisiblePixels,
                "Native PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + "/" + totalPixels + ". Path: " + pngPath);
        }

        private static void ThrowBaselineChanged(string baselineName, string expectedPath, string actualPath, VisualRasterComparison? comparison, VisioPremiumBaselineContext context) {
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
                  "channel tolerance: " + comparison.ChannelTolerance + "; " +
                  "MAE: " + comparison.MeanAbsoluteError.ToString("0.###", CultureInfo.InvariantCulture) + "/" +
                    comparison.MaximumMeanAbsoluteError.ToString("0.###", CultureInfo.InvariantCulture) + "; " +
                  "RMSE: " + comparison.RootMeanSquareError.ToString("0.###", CultureInfo.InvariantCulture) + "/" +
                    comparison.MaximumRootMeanSquareError.ToString("0.###", CultureInfo.InvariantCulture) + "; " +
                  "luminance MAE: " + comparison.MeanLuminanceError.ToString("0.###", CultureInfo.InvariantCulture) + "/" +
                    comparison.MaximumMeanLuminanceError.ToString("0.###", CultureInfo.InvariantCulture) + ". ";
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
            string expectedPath = Path.Combine(VisualBaselineTestSupport.GetTestsProjectRoot(), "Visio", "VisualBaselines", baselineName);
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
            return VisualBaselineTestSupport.CreateArtifactDirectory("OfficeIMO.VisioPremiumBaselines");
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

        private static string NormalizeText(string text) =>
            VisualBaselineTestSupport.NormalizeTextWithTrailingNewLine(text);

        private static bool IsNativePngBaseline(string baselineName) =>
            baselineName.IndexOf("-native-", StringComparison.OrdinalIgnoreCase) >= 0 &&
            baselineName.EndsWith(".png", StringComparison.OrdinalIgnoreCase);

        private static string ResolveExpectedBaselineName(string baselineName) =>
            RuntimeInformation.IsOSPlatform(OSPlatform.Linux) &&
            string.Equals(baselineName, PrintAuditTrailNativePngBaseline, StringComparison.OrdinalIgnoreCase)
                ? LinuxPrintAuditTrailNativePngBaseline
                : baselineName;

        private static VisualRasterComparison CompareRasterImages(byte[] expectedPng, byte[] actualPng, bool allowNativeVariance = false) {
            int channelTolerance = VisualBaselineTestSupport.ReadNonNegativeInt("OFFICEIMO_VISIO_PREMIUM_BASELINE_PIXEL_TOLERANCE", 0);
            int allowedDifferentPixels = VisualBaselineTestSupport.ReadNonNegativeInt("OFFICEIMO_VISIO_PREMIUM_BASELINE_ALLOWED_DIFF_PIXELS", 1);
            OfficeRasterImage expected = VisualBaselineTestSupport.DecodePng(expectedPng, "Expected Visio premium baseline is not a supported PNG file.");
            OfficeRasterImage actual = VisualBaselineTestSupport.DecodePng(actualPng, "Actual Visio premium output is not a supported PNG file.");
            if (allowNativeVariance && expected.Width == actual.Width && expected.Height == actual.Height) {
                // Native PNG text rasterization varies between platform font stacks. Exact SVG and
                // structural baselines continue to protect the content, geometry, and styling contract.
                int defaultAllowedDifferentPixels = Math.Max(1, expected.Width * expected.Height / 50);
                allowedDifferentPixels = VisualBaselineTestSupport.ReadNonNegativeInt("OFFICEIMO_VISIO_PREMIUM_NATIVE_BASELINE_ALLOWED_DIFF_PIXELS", defaultAllowedDifferentPixels);
            }

            double defaultMaximumMeanAbsoluteError = allowNativeVariance ? 2.5D : 0D;
            double defaultMaximumRootMeanSquareError = allowNativeVariance ? 18D : 0D;
            double defaultMaximumMeanLuminanceError = allowNativeVariance ? 3.5D : 0D;
            return VisualBaselineTestSupport.CompareRasterImages(
                expected,
                actual,
                channelTolerance,
                allowedDifferentPixels,
                VisualBaselineTestSupport.ReadNonNegativeDouble("OFFICEIMO_VISIO_PREMIUM_BASELINE_MAX_MAE", defaultMaximumMeanAbsoluteError),
                VisualBaselineTestSupport.ReadNonNegativeDouble("OFFICEIMO_VISIO_PREMIUM_BASELINE_MAX_RMSE", defaultMaximumRootMeanSquareError),
                VisualBaselineTestSupport.ReadNonNegativeDouble("OFFICEIMO_VISIO_PREMIUM_BASELINE_MAX_LUMINANCE_MAE", defaultMaximumMeanLuminanceError));
        }

        private static VisualRasterComparison CompareRasterImages(byte[] expectedPng, byte[] actualPng, int channelTolerance, int allowedDifferentPixels) =>
            VisualBaselineTestSupport.CompareRasterImages(expectedPng, actualPng, channelTolerance, allowedDifferentPixels);

        private static string CanonicalizeSvg(string path) {
            string svg = File.ReadAllText(path)
                .Replace("\r\n", "\n")
                .Replace("\r", "\n");

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
            svg = Regex.Replace(svg, @"<title>.*?</title>\s*", string.Empty, RegexOptions.Singleline | RegexOptions.CultureInvariant);
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
            if (normalized.IndexOf("stroke:none", StringComparison.Ordinal) >= 0) {
                normalized = Regex.Replace(normalized, @";?stroke-width:[^;]+", string.Empty, RegexOptions.CultureInvariant);
            }

            return normalized.Trim(';');
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

        private static string FormatGalleryIssues(VisioGalleryResult result) {
            IEnumerable<string> issues = result.PackageIssues
                .Concat(result.QualityIssues.Select(issue => issue.ToString()));
            return result.Name + Environment.NewLine + string.Join(Environment.NewLine, issues);
        }

        private static bool IsRequired() =>
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_REQUIRE_VISIO_PREMIUM_BASELINES"), "1", StringComparison.Ordinal);

        private static bool IsDesktopBaselineRunRequested() =>
            IsRequired() ||
            IsBaselineUpdateRequested() ||
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_RUN_VISIO_PREMIUM_DESKTOP_BASELINES"), "1", StringComparison.Ordinal);

        private static bool IsBaselineUpdateRequested() =>
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_VISIO_PREMIUM_BASELINES"), "1", StringComparison.Ordinal);

        private static bool IsStructuralBaselineUpdateRequested() =>
            IsBaselineUpdateRequested() ||
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_VISIO_PREMIUM_STRUCTURAL_BASELINES"), "1", StringComparison.Ordinal);

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
