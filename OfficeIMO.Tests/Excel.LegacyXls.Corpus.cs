using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_Corpus_Fixtures_MatchApprovedImportReports() {
            string corpusDirectory = Path.Combine(GetTestsProjectRoot(), "Documents", "LegacyXlsCorpus");
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

            bool updateBaselines = IsLegacyXlsCorpusBaselineUpdateRequested();
            var missingBaselines = new List<string>();
            foreach (string workbookPath in workbookPaths) {
                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                    ReportUnsupportedRecords = true
                });
                string actual = NormalizeBaselineText(result.ImportReport.ToMarkdown());
                string baselinePath = Path.ChangeExtension(workbookPath, ".import-report.md");

                if (updateBaselines) {
                    File.WriteAllText(baselinePath, actual, Encoding.UTF8);
                    continue;
                }

                if (!File.Exists(baselinePath)) {
                    missingBaselines.Add(GetRelativePath(corpusDirectory, baselinePath));
                    continue;
                }

                string expected = NormalizeBaselineText(File.ReadAllText(baselinePath, Encoding.UTF8));
                Assert.Equal(expected, actual);
            }

            Assert.True(
                missingBaselines.Count == 0,
                "Missing legacy XLS corpus baselines. Run with OFFICEIMO_UPDATE_LEGACY_XLS_CORPUS_BASELINES=1 to create: "
                    + string.Join(", ", missingBaselines));
        }

        private static bool IsLegacyXlsCorpusBaselineUpdateRequested() {
            string? value = Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_LEGACY_XLS_CORPUS_BASELINES");
            return string.Equals(value, "1", StringComparison.Ordinal)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeBaselineText(string text) {
            return text.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd() + "\n";
        }

        private static string GetTestsProjectRoot() {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null) {
                if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.Tests.csproj"))) {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            return AppContext.BaseDirectory;
        }

        private static string GetRelativePath(string relativeTo, string path) {
            string root = EnsureTrailingDirectorySeparator(Path.GetFullPath(relativeTo));
            string target = Path.GetFullPath(path);
            Uri rootUri = new Uri(root);
            Uri targetUri = new Uri(target);
            string relative = Uri.UnescapeDataString(rootUri.MakeRelativeUri(targetUri).ToString());
            return relative.Replace('/', Path.DirectorySeparatorChar);
        }

        private static string EnsureTrailingDirectorySeparator(string path) {
            char separator = Path.DirectorySeparatorChar;
            char alternateSeparator = Path.AltDirectorySeparatorChar;
            if (path.Length == 0 || path[path.Length - 1] == separator || path[path.Length - 1] == alternateSeparator) {
                return path;
            }

            return path + separator;
        }
    }
}
