using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    [Trait("Category", "ExcelImageVisualGate")]
    public class ExcelImageExportVisualFidelityGateTests {
        private static readonly IReadOnlyList<ExcelImageBaselineFidelityRecord> Baselines = new[] {
            Clean("officeimo-excel-image-aligned-shape-text"),
            Clean("officeimo-excel-image-chart-axis-labels"),
            Clean("officeimo-excel-image-chart-data-label-boxes"),
            Tracked("officeimo-excel-image-clipped-image", "Cell text intentionally suppressed by later drawing layer.", ExcelImageExportDiagnosticCodes.CellTextOccludedByDrawing),
            Tracked("officeimo-excel-image-comment-body", "Comment body is rendered as a dependency-free callout approximation.", ExcelImageExportDiagnosticCodes.CellCommentBodyApproximation),
            Tracked("officeimo-excel-image-conditional-formatting", "Icon artwork and threshold parity remain approximate.", ExcelImageExportDiagnosticCodes.ConditionalIconSetApproximation),
            Clean("officeimo-excel-image-cropped-image"),
            Clean("officeimo-excel-image-drawing-object"),
            Tracked(
                "officeimo-excel-image-expanded-icon-sets",
                "Expanded icon-set visuals remain deterministic approximations.",
                ExcelImageExportDiagnosticCodes.ConditionalIconSetApproximation,
                ExcelImageExportDiagnosticCodes.ConditionalIconSetApproximation,
                ExcelImageExportDiagnosticCodes.ConditionalIconSetApproximation,
                ExcelImageExportDiagnosticCodes.ConditionalIconSetApproximation),
            Clean("officeimo-excel-image-header-footer-images"),
            Clean("officeimo-excel-image-page-layout"),
            Tracked(
                "officeimo-excel-image-pattern-fills",
                "Pattern fills are deterministic hatch approximations.",
                ExcelImageExportDiagnosticCodes.FillPatternApproximation,
                ExcelImageExportDiagnosticCodes.FillPatternApproximation,
                ExcelImageExportDiagnosticCodes.FillPatternApproximation,
                ExcelImageExportDiagnosticCodes.FillPatternApproximation,
                ExcelImageExportDiagnosticCodes.FillPatternApproximation,
                ExcelImageExportDiagnosticCodes.FillPatternApproximation,
                ExcelImageExportDiagnosticCodes.FillPatternApproximation),
            Tracked(
                "officeimo-excel-image-premium-range",
                "Premium range renders the comment body as a dependency-free callout approximation and reports the unavailable Aptos chart font before using the managed stroke fallback.",
                ExcelImageExportDiagnosticCodes.CellCommentBodyApproximation,
                OfficeImageExportDiagnosticCodes.FontSubstituted),
            Tracked("officeimo-excel-image-rich-text", "Rich text still has clipped and rotated-text fidelity gaps.", ExcelImageExportDiagnosticCodes.CellTextClipped, ExcelImageExportDiagnosticCodes.CellTextRotationApproximation),
            Clean("officeimo-excel-image-rotated-image"),
            Clean("officeimo-excel-image-rotated-preset-drawing-object"),
            Tracked("officeimo-excel-image-rotated-shape-text", "Rotated shape text routes through Drawing but text-box metrics are not Excel-exact.", ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation),
            Tracked("officeimo-excel-image-sparklines", "Sparkline hidden/empty/date-axis behavior remains approximate.", ExcelImageExportDiagnosticCodes.SparklineRenderingApproximation, ExcelImageExportDiagnosticCodes.SparklineRenderingApproximation, ExcelImageExportDiagnosticCodes.SparklineRenderingApproximation),
            Tracked("officeimo-excel-image-stacked-text", "Stacked text is readable but baseline metrics remain approximate.", ExcelImageExportDiagnosticCodes.CellTextRotationApproximation, ExcelImageExportDiagnosticCodes.CellTextRotationApproximation, ExcelImageExportDiagnosticCodes.CellTextRotationApproximation),
            Clean("officeimo-excel-image-text-spill"),
            Clean("officeimo-excel-image-transformed-image"),
            Clean("officeimo-excel-image-two-cell-image"),
            Clean("officeimo-excel-image-vertical-shape-text"),
            Tracked("officeimo-excel-image-vertical270-shape-text", "Vertical270 shape text routes through Drawing rotation but metrics are not Excel-exact.", ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation),
        };

        [Fact]
        public void ExcelImageVisualFidelityManifestCoversEveryApprovedBaseline() {
            string baselineDirectory = BaselineDirectory;
            string[] approvedPngBaselines = Directory
                .GetFiles(baselineDirectory, "*.png", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFileNameWithoutExtension)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Select(name => name!)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();

            string[] manifestBaselines = Baselines
                .Select(item => item.Name)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();

            Assert.Equal(approvedPngBaselines, manifestBaselines);
        }

        [Fact]
        public void ExcelImageVisualFidelityManifestTracksEveryDiagnosticsBaseline() {
            string baselineDirectory = BaselineDirectory;
            string[] diagnosticsBaselines = Directory
                .GetFiles(baselineDirectory, "*.diagnostics.txt", SearchOption.TopDirectoryOnly)
                .Select(path => TrimDiagnosticsBaselineSuffix(Path.GetFileName(path)))
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();

            string[] trackedBaselines = Baselines
                .Where(item => item.ExpectedDiagnosticCodes.Count > 0)
                .Select(item => item.Name)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();

            Assert.Equal(diagnosticsBaselines, trackedBaselines);
        }

        [Fact]
        public void ExcelImageVisualFidelityManifestPinsKnownDiagnosticsAndGaps() {
            foreach (ExcelImageBaselineFidelityRecord baseline in Baselines) {
                Assert.NotNull(baseline.ReviewNote);

                string diagnosticsPath = Path.Combine(BaselineDirectory, baseline.Name + ".diagnostics.txt");
                if (baseline.ExpectedDiagnosticCodes.Count == 0) {
                    Assert.False(File.Exists(diagnosticsPath), baseline.Name + " is marked clean but carries a diagnostics baseline.");
                    continue;
                }

                Assert.True(File.Exists(diagnosticsPath), baseline.Name + " is marked tracked but has no diagnostics baseline.");
                IReadOnlyList<ParsedDiagnostic> diagnostics = ParseDiagnosticsBaseline(diagnosticsPath);

                Assert.DoesNotContain(diagnostics, diagnostic => string.Equals(diagnostic.Severity, "Error", StringComparison.Ordinal));
                string[] actualCodes = diagnostics
                    .Select(item => item.Code)
                    .OrderBy(code => code, StringComparer.Ordinal)
                    .ToArray();
                string[] expectedCodes = baseline.ExpectedDiagnosticCodes
                    .OrderBy(code => code, StringComparer.Ordinal)
                    .ToArray();

                Assert.Equal(expectedCodes, actualCodes);
                Assert.All(diagnostics, diagnostic => Assert.False(string.IsNullOrWhiteSpace(diagnostic.Source), baseline.Name + " diagnostic has no source reference."));
            }
        }

        [Fact]
        public void ExcelImageVisualFidelityManifestTracksPremiumRangeCommentBodyApproximation() {
            ExcelImageBaselineFidelityRecord premiumRange = Assert.Single(Baselines, item => item.Name == "officeimo-excel-image-premium-range");

            Assert.Equal(ExcelImageBaselineFidelity.TrackedApproximation, premiumRange.Fidelity);
            Assert.Contains(ExcelImageExportDiagnosticCodes.CellCommentBodyApproximation, premiumRange.ExpectedDiagnosticCodes);
            Assert.Contains(OfficeImageExportDiagnosticCodes.FontSubstituted, premiumRange.ExpectedDiagnosticCodes);
            Assert.DoesNotContain(ExcelImageExportDiagnosticCodes.CellCommentUnsupported, premiumRange.ExpectedDiagnosticCodes);
        }

        private static ExcelImageBaselineFidelityRecord Clean(string name) =>
            new(name, ExcelImageBaselineFidelity.CleanApprovedBaseline, "Approved baseline has no current exporter diagnostics.", Array.Empty<string>());

        private static ExcelImageBaselineFidelityRecord Tracked(string name, string reviewNote, params string[] expectedDiagnosticCodes) =>
            new(name, ExcelImageBaselineFidelity.TrackedApproximation, reviewNote, expectedDiagnosticCodes);

        private static IReadOnlyList<ParsedDiagnostic> ParseDiagnosticsBaseline(string path) {
            var diagnostics = new List<ParsedDiagnostic>();
            foreach (string line in File.ReadAllLines(path)) {
                if (string.IsNullOrWhiteSpace(line)) {
                    continue;
                }

                string[] parts = line.Split('|');
                Assert.True(parts.Length >= 4, "Diagnostics baseline line must contain severity, code, source, and message: " + line);
                diagnostics.Add(new ParsedDiagnostic(parts[0], parts[1], parts[2]));
            }

            return diagnostics;
        }

        private static string TrimDiagnosticsBaselineSuffix(string fileName) {
            const string suffix = ".diagnostics.txt";
            return fileName.EndsWith(suffix, StringComparison.Ordinal)
                ? fileName.Substring(0, fileName.Length - suffix.Length)
                : Path.GetFileNameWithoutExtension(fileName);
        }

        private static string BaselineDirectory =>
            Path.Combine(VisualBaselineTestSupport.GetTestsProjectRoot(), "Excel", "VisualBaselines");

        private enum ExcelImageBaselineFidelity {
            CleanApprovedBaseline,
            TrackedApproximation
        }

        private sealed record ExcelImageBaselineFidelityRecord(
            string Name,
            ExcelImageBaselineFidelity Fidelity,
            string ReviewNote,
            IReadOnlyList<string> ExpectedDiagnosticCodes);

        private sealed record ParsedDiagnostic(string Severity, string Code, string Source);
    }
}
