using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
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

        [Fact]
        public void LegacyXls_Corpus_ExcelExternalLinks_PreserveFormulaAndCacheModel() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "external-links.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            Assert.Equal(1, result.ImportReport.ExternalReferencesByTarget["external-source.xls"]);

            LegacyXlsExternalReference externalReference = Assert.Single(
                result.Workbook.ExternalReferences,
                reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
            Assert.Equal("\u0001external-source.xls", externalReference.Target);
            Assert.Equal(new[] { "Data" }, externalReference.SheetNames);

            LegacyXlsExternalCellCache cache = Assert.Single(externalReference.CachedCellCaches);
            Assert.True(cache.LinkValid);
            Assert.Equal("Data", cache.SheetName);
            Assert.Equal("R1C1:R3C1", cache.CellRange);
            Assert.Collection(cache.Cells,
                cell => {
                    Assert.Equal(LegacyXlsCellValueKind.Number, cell.Kind);
                    Assert.Equal(125d, cell.Value);
                },
                cell => {
                    Assert.Equal(LegacyXlsCellValueKind.Number, cell.Kind);
                    Assert.Equal(80d, cell.Value);
                },
                cell => {
                    Assert.Equal(LegacyXlsCellValueKind.Number, cell.Kind);
                    Assert.Equal(210d, cell.Value);
                });

            LegacyXlsWorksheet sheet = Assert.Single(result.Workbook.Worksheets);
            AssertCorpusFormula(sheet, 1, 2, 125d, "'[external-source.xls]Data'!$B$2");
            AssertCorpusFormula(sheet, 2, 2, 80d, "'[external-source.xls]Data'!$B$3");
            AssertCorpusFormula(sheet, 3, 2, 45d, "B1-B2");
            AssertCorpusFormula(sheet, 5, 2, 415d, "SUM('[external-source.xls]Data'!$B$2:$B$4)");
        }

        [Fact]
        public void LegacyXls_Corpus_ExcelObjects_PreserveDrawingObjectSubrecordModel() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "objects.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.True(result.ImportReport.DrawingRecordsByObjectTypeName["Picture"] >= 1);
            Assert.True(result.ImportReport.DrawingRecordsByObjectTypeName["Button"] >= 1);
            Assert.True(result.ImportReport.DrawingRecordsByObjectTypeName["Checkbox"] >= 1);
            Assert.True(result.ImportReport.DrawingRecordsByObjectTypeName["DropdownList"] >= 1);
            Assert.True(result.ImportReport.DrawingObjectSubRecordsByName["FtCmo"] >= 1);
            Assert.True(result.ImportReport.DrawingObjectSubRecordsByName["FtEnd"] >= 1);
            Assert.True(result.ImportReport.DrawingObjectSubRecordsByName["FtCblsData"] >= 1);
            Assert.True(result.ImportReport.DrawingObjectSubRecordsByName["FtLbsData"] >= 1);
            Assert.True(result.ImportReport.DrawingObjectSubRecordsByCompleteness["Complete"] >= 1);
            Assert.True(result.ImportReport.DrawingObjectSubRecordsByCompleteness["Truncated"] >= 1);
            Assert.Equal(1, result.ImportReport.DrawingBlipStoreEntriesByEmbeddedRecordType["OfficeArtBlipPNG"]);
            Assert.Contains(result.Workbook.DrawingRecords, record => record.ObjectTypeName == "Picture" && record.HasObjectSubRecords);
            Assert.Contains(result.Workbook.DrawingRecords, record => record.ObjectTypeName == "Button" && record.HasObjectSubRecords);
            Assert.Contains(result.Workbook.DrawingRecords, record => record.ObjectTypeName == "Checkbox" && record.ObjectSubRecords.Any(subRecord => subRecord.SubRecordName == "FtCblsData"));
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

        private static void AssertCorpusFormula(LegacyXlsWorksheet sheet, int row, int column, object expectedValue, string expectedFormulaText) {
            LegacyXlsCell? cell = sheet.Cells.SingleOrDefault(candidate => candidate.Row == row && candidate.Column == column);
            Assert.True(
                cell != null,
                "Expected formula cell was not found. Parsed formula cells: "
                    + string.Join(", ", sheet.Cells
                        .Where(candidate => candidate.IsFormula)
                        .Select(candidate => $"R{candidate.Row}C{candidate.Column}={candidate.FormulaText}")));
            Assert.True(cell.IsFormula);
            Assert.Equal(expectedValue, cell.Value);
            Assert.Equal(expectedFormulaText, cell.FormulaText);
        }
    }
}
