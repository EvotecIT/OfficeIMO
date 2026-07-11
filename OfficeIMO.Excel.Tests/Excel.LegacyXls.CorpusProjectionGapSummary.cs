using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_Corpus_ProjectionGapSummary_MatchesApprovedBaseline() {
            string corpusDirectory = Path.Combine(GetProjectionGapTestsProjectRoot(), "Documents", "LegacyXlsCorpus");
            AssertLegacyXlsProjectionGapSummaryBaseline(corpusDirectory);
        }

        [Fact]
        public void LegacyXls_DiagnosticCorpus_ProjectionGapSummary_MatchesApprovedBaseline() {
            string corpusDirectory = Path.Combine(GetProjectionGapTestsProjectRoot(), "Documents", "LegacyXlsDiagnosticCorpus");
            AssertLegacyXlsProjectionGapSummaryBaseline(corpusDirectory);
        }

        private static void AssertLegacyXlsProjectionGapSummaryBaseline(string corpusDirectory) {
            if (!Directory.Exists(corpusDirectory)) {
                return;
            }

            string[] workbookPaths = Directory.GetFiles(corpusDirectory, "*.xls", SearchOption.AllDirectories)
                .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (workbookPaths.Length == 0) {
                return;
            }

            string actual = NormalizeProjectionGapSummaryText(BuildLegacyXlsProjectionGapSummary(corpusDirectory, workbookPaths));
            string baselinePath = Path.Combine(corpusDirectory, "projection-gap-summary.md");

            if (IsLegacyXlsProjectionGapBaselineUpdateRequested()) {
                File.WriteAllText(baselinePath, actual, Encoding.UTF8);
                return;
            }

            Assert.True(
                File.Exists(baselinePath),
                "Missing legacy XLS projection-gap summary. Run with OFFICEIMO_UPDATE_LEGACY_XLS_CORPUS_BASELINES=1 to create: "
                    + GetProjectionGapRelativePath(corpusDirectory, baselinePath));

            string expected = NormalizeProjectionGapSummaryText(File.ReadAllText(baselinePath, Encoding.UTF8));
            Assert.Equal(expected, actual);
        }

        private static string BuildLegacyXlsProjectionGapSummary(string corpusDirectory, IReadOnlyList<string> workbookPaths) {
            var fixtureRows = new List<LegacyXlsProjectionGapFixtureRow>(workbookPaths.Count);
            var gapsByKind = new SortedDictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var gapsByDetail = new SortedDictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int totalProjectionGaps = 0;
            int totalErrors = 0;
            int totalWarnings = 0;

            foreach (string workbookPath in workbookPaths) {
                LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookPath, new LegacyXlsImportOptions {
                    ReportUnsupportedContent = true
                });
                LegacyXlsImportReport report = workbook.CreateImportReport();
                string fixture = GetProjectionGapRelativePath(corpusDirectory, workbookPath).Replace('\\', '/');

                fixtureRows.Add(new LegacyXlsProjectionGapFixtureRow(
                    fixture,
                    report.UnsupportedProjectionGapCount,
                    report.ErrorCount,
                    report.WarningCount));

                totalProjectionGaps += report.UnsupportedProjectionGapCount;
                totalErrors += report.ErrorCount;
                totalWarnings += report.WarningCount;
                AddCounts(gapsByKind, report.UnsupportedProjectionGapsByKind.ToDictionary(
                    entry => entry.Key.ToString(),
                    entry => entry.Value,
                    StringComparer.OrdinalIgnoreCase));
                AddCounts(gapsByDetail, report.UnsupportedProjectionGapsByDetail);
            }

            var builder = new StringBuilder();
            builder.AppendLine("# Legacy XLS Corpus Projection Gap Summary");
            builder.AppendLine();
            builder.AppendLine("Corpus: " + Path.GetFileName(corpusDirectory));
            builder.AppendLine("Fixtures: " + fixtureRows.Count.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Unsupported projection gaps: " + totalProjectionGaps.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Fixtures with projection gaps: " + fixtureRows.Count(row => row.ProjectionGapCount > 0).ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Errors: " + totalErrors.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Warnings: " + totalWarnings.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine();
            AppendFixtureRows(builder, fixtureRows);
            AppendSummaryDictionary(builder, "Projection Gaps By Kind", gapsByKind);
            AppendSummaryDictionary(builder, "Projection Gap Details", gapsByDetail);
            return builder.ToString();
        }

        private static void AppendFixtureRows(StringBuilder builder, IEnumerable<LegacyXlsProjectionGapFixtureRow> rows) {
            builder.AppendLine("## Fixtures");
            builder.AppendLine();
            builder.AppendLine("| Fixture | Projection gaps | Errors | Warnings |");
            builder.AppendLine("| --- | --- | --- | --- |");
            foreach (LegacyXlsProjectionGapFixtureRow row in rows.OrderBy(row => row.Fixture, StringComparer.OrdinalIgnoreCase)) {
                builder.Append("| ");
                builder.Append(EscapeProjectionGapMarkdownCell(row.Fixture));
                builder.Append(" | ");
                builder.Append(row.ProjectionGapCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" | ");
                builder.Append(row.ErrorCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" | ");
                builder.Append(row.WarningCount.ToString(CultureInfo.InvariantCulture));
                builder.AppendLine(" |");
            }
        }

        private static void AppendSummaryDictionary(StringBuilder builder, string title, IReadOnlyDictionary<string, int> values) {
            builder.AppendLine();
            builder.AppendLine("## " + title);
            builder.AppendLine();
            if (values.Count == 0) {
                builder.AppendLine("(none)");
                return;
            }

            builder.AppendLine("| Key | Count |");
            builder.AppendLine("| --- | --- |");
            foreach (KeyValuePair<string, int> entry in values.OrderBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)) {
                builder.Append("| ");
                builder.Append(EscapeProjectionGapMarkdownCell(entry.Key));
                builder.Append(" | ");
                builder.Append(entry.Value.ToString(CultureInfo.InvariantCulture));
                builder.AppendLine(" |");
            }
        }

        private static void AddCounts(IDictionary<string, int> target, IReadOnlyDictionary<string, int> source) {
            foreach (KeyValuePair<string, int> entry in source) {
                target.TryGetValue(entry.Key, out int current);
                target[entry.Key] = current + entry.Value;
            }
        }

        private static bool IsLegacyXlsProjectionGapBaselineUpdateRequested() {
            string? value = Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_LEGACY_XLS_CORPUS_BASELINES");
            return string.Equals(value, "1", StringComparison.Ordinal)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeProjectionGapSummaryText(string text) {
            return text.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd() + "\n";
        }

        private static string EscapeProjectionGapMarkdownCell(string value) {
            return value.Replace("\\", "\\\\").Replace("|", "\\|");
        }

        private static string GetProjectionGapTestsProjectRoot() {
            bool updateBaselines = IsLegacyXlsProjectionGapBaselineUpdateRequested();
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null) {
                string legacyTestRoot = Path.Combine(directory.FullName, "OfficeIMO.TestAssets");
                if (updateBaselines && Directory.Exists(Path.Combine(legacyTestRoot, "Documents", "LegacyXlsCorpus"))) {
                    return legacyTestRoot;
                }

                if (Directory.Exists(Path.Combine(directory.FullName, "Documents", "LegacyXlsCorpus"))) {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            throw new DirectoryNotFoundException("Unable to locate OfficeIMO legacy XLS fixture root from " + AppContext.BaseDirectory);
        }

        private static string GetProjectionGapRelativePath(string relativeTo, string path) {
            string root = EnsureProjectionGapTrailingDirectorySeparator(Path.GetFullPath(relativeTo));
            string target = Path.GetFullPath(path);
            Uri rootUri = new Uri(root);
            Uri targetUri = new Uri(target);
            string relative = Uri.UnescapeDataString(rootUri.MakeRelativeUri(targetUri).ToString());
            return relative.Replace('/', Path.DirectorySeparatorChar);
        }

        private static string EnsureProjectionGapTrailingDirectorySeparator(string path) {
            char separator = Path.DirectorySeparatorChar;
            char alternateSeparator = Path.AltDirectorySeparatorChar;
            if (path.Length == 0 || path[path.Length - 1] == separator || path[path.Length - 1] == alternateSeparator) {
                return path;
            }

            return path + separator;
        }

        private sealed record LegacyXlsProjectionGapFixtureRow(
            string Fixture,
            int ProjectionGapCount,
            int ErrorCount,
            int WarningCount);
    }
}
