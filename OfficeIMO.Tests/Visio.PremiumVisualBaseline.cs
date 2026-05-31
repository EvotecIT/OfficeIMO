using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioPremiumVisualBaselineTests {
        private static readonly IReadOnlyDictionary<string, string> BaselinePrefixes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["Premium Cloud Architecture"] = "officeimo-visio-premium-cloud-architecture",
            ["Premium Network Segmentation"] = "officeimo-visio-premium-network-segmentation",
            ["Premium Executive Dependencies"] = "officeimo-visio-premium-executive-dependencies",
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
                        AssertBaseline(Path.GetFileName(actualPath), actualPath);
                    }
                }
            } finally {
                TryDeleteDirectory(workDirectory);
            }
        }

        private static string GetBaselinePrefix(string name) {
            if (BaselinePrefixes.TryGetValue(name, out string? prefix)) {
                return prefix;
            }

            throw new InvalidOperationException("No premium Visio visual baseline prefix is registered for '" + name + "'.");
        }

        private static void AssertBaseline(string baselineName, string actualPath) {
            string expectedPath = Path.Combine(GetTestsProjectRoot(), "Visio", "VisualBaselines", baselineName);
            if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_VISIO_PREMIUM_BASELINES"), "1", StringComparison.Ordinal)) {
                Directory.CreateDirectory(Path.GetDirectoryName(expectedPath)!);
                File.Copy(actualPath, expectedPath, overwrite: true);
                return;
            }

            if (!File.Exists(expectedPath)) {
                throw new FileNotFoundException(
                    "Premium Visio visual baseline missing. Set OFFICEIMO_UPDATE_VISIO_PREMIUM_BASELINES=1 and re-run this test to generate it.",
                    expectedPath);
            }

            bool matches = string.Equals(Path.GetExtension(expectedPath), ".svg", StringComparison.OrdinalIgnoreCase)
                ? string.Equals(CanonicalizeSvg(expectedPath), CanonicalizeSvg(actualPath), StringComparison.Ordinal)
                : File.ReadAllBytes(expectedPath).AsSpan().SequenceEqual(File.ReadAllBytes(actualPath));
            if (matches) {
                return;
            }

            string artifactDirectory = Path.Combine(
                Path.GetTempPath(),
                "OfficeIMO.VisioPremiumBaselines",
                DateTime.UtcNow.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(artifactDirectory);

            File.Copy(expectedPath, Path.Combine(artifactDirectory, "expected-" + Path.GetFileName(expectedPath)), overwrite: true);
            File.Copy(actualPath, Path.Combine(artifactDirectory, "actual-" + Path.GetFileName(actualPath)), overwrite: true);

            FileInfo expectedInfo = new(expectedPath);
            FileInfo actualInfo = new(actualPath);
            throw new Xunit.Sdk.XunitException(
                "Premium Visio visual baseline changed for '" + baselineName + "'. " +
                "Expected bytes: " + expectedInfo.Length + "; actual bytes: " + actualInfo.Length + ". " +
                "Artifacts: " + artifactDirectory + ".");
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

        private static string FormatGalleryIssues(VisioGalleryResult result) {
            IEnumerable<string> issues = result.PackageIssues
                .Concat(result.QualityIssues.Select(issue => issue.ToString()))
                .Concat(result.DesktopValidation?.Issues ?? Array.Empty<string>());
            return result.Name + Environment.NewLine + string.Join(Environment.NewLine, issues);
        }

        private static bool IsRequired() =>
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_REQUIRE_VISIO_PREMIUM_BASELINES"), "1", StringComparison.Ordinal);

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
