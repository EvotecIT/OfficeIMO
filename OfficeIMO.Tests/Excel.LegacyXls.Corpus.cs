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
            AssertLegacyXlsCorpusBaselines(corpusDirectory);
        }

        [Fact]
        public void LegacyXls_DiagnosticCorpus_Fixtures_MatchApprovedImportReports() {
            string corpusDirectory = Path.Combine(GetTestsProjectRoot(), "Documents", "LegacyXlsDiagnosticCorpus");
            AssertLegacyXlsCorpusBaselines(corpusDirectory);
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
            Assert.Equal(1, result.ImportReport.ExternalSheetNamesByTarget["external-source.xls!Data"]);
            Assert.Equal(1, result.ImportReport.ExternalCellCachesByTarget["external-source.xls"]);
            Assert.Equal(1, result.ImportReport.ExternalCellCachesByTargetAndSheetName["external-source.xls!Data"]);
            Assert.Equal(1, result.ImportReport.ExternalCellCachesByTargetAndCellRange["external-source.xls!R1C1:R3C1"]);
            Assert.Equal(3, result.ImportReport.ExternalCachedCellsByTargetSheetAndValueKind["external-source.xls!Data|Number"]);

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
            Assert.Equal(1, result.ImportReport.DrawingBlipStoreEntriesByLocation["(workbook)"]);
            Assert.Equal(1, result.ImportReport.DrawingBlipStoreEntriesByTypeAndLocation["(workbook)|Png"]);
            Assert.Equal(1, result.ImportReport.DrawingShapePropertiesByName["pib"]);
            Assert.Equal(1, result.ImportReport.DrawingShapePropertiesByName["BlipBooleanProperties"]);
            Assert.Equal(1, result.ImportReport.DrawingShapePropertiesByName["ShapeBooleanProperties"]);
            Assert.Equal(2, result.ImportReport.DrawingShapePropertiesByName["wzName"]);
            Assert.Equal(2, result.ImportReport.DrawingShapePropertiesByGroup["Blip"]);
            Assert.Equal(2, result.ImportReport.DrawingShapeBlipPropertiesByLocation["Objects"]);
            Assert.Equal(1, result.ImportReport.DrawingShapeBlipPropertiesByNameAndValue["pib;Value:0x00000001"]);
            Assert.Equal(1, result.ImportReport.DrawingShapeBlipPropertiesByNameAndValue["BlipBooleanProperties;Value:0x00060000"]);
            Assert.Equal(1, result.ImportReport.DrawingPictureBlipReferencesByLocation["Objects"]);
            Assert.Equal(1, result.ImportReport.DrawingPictureBlipReferencesByValue["BlipId:1"]);
            Assert.Equal(1, result.ImportReport.DrawingShapePropertiesByGroup["Shape"]);
            Assert.False(result.ImportReport.DrawingShapePropertiesByGroup.ContainsKey("Unknown"));
            Assert.Equal(1, result.ImportReport.DrawingShapeComplexPropertiesByText["wzName:Chart 5"]);
            Assert.Equal(1, result.ImportReport.DrawingShapeComplexPropertiesByText["wzName:Picture 4"]);
            Assert.Contains(result.Workbook.DrawingRecords.SelectMany(record => record.ShapeProperties),
                property => property.PropertyName == "wzName" && property.ComplexText == "Chart 5");
            Assert.Contains(result.Workbook.DrawingRecords.SelectMany(record => record.ShapeProperties),
                property => property.PropertyName == "wzName" && property.ComplexText == "Picture 4");
            Assert.Contains(result.Workbook.DrawingRecords, record => record.ObjectTypeName == "Picture" && record.HasObjectSubRecords);
            Assert.Contains(result.Workbook.DrawingRecords, record => record.ObjectTypeName == "Button" && record.HasObjectSubRecords);
            Assert.Contains(result.Workbook.DrawingRecords, record => record.ObjectTypeName == "Checkbox" && record.ObjectSubRecords.Any(subRecord => subRecord.SubRecordName == "FtCblsData"));
        }

        [Fact]
        public void LegacyXls_Corpus_FormulaStress_ProjectsSupportedFormulaTokens() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "formula-stress.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            Assert.Empty(result.ImportReport.FormulaTokenBlockers);
            Assert.Contains("ROUND", result.ImportReport.FormulaFunctionsByName.Keys);
            Assert.Contains("IF", result.ImportReport.FormulaFunctionsByName.Keys);
            Assert.Contains("If", result.ImportReport.FormulaAttributesByName.Keys);
            Assert.Contains("Sum", result.ImportReport.FormulaAttributesByName.Keys);

            foreach (string tokenName in new[] {
                "PtgAdd",
                "PtgArea",
                "PtgArray",
                "PtgAttr",
                "PtgConcat",
                "PtgDiv",
                "PtgFunc",
                "PtgFuncVar",
                "PtgGt",
                "PtgLe",
                "PtgNe",
                "PtgPercent",
                "PtgPower",
                "PtgStr",
                "PtgUminus",
                "PtgUplus"
            }) {
                Assert.Contains(tokenName, result.ImportReport.FormulaTokensByName.Keys);
            }

            LegacyXlsWorksheet sheet = Assert.Single(result.Workbook.Worksheets);
            AssertCorpusFormula(sheet, 1, 2, 100d, "A1^2");
            AssertCorpusFormula(sheet, 2, 2, "North-Q1", "A4&\"-\"&A5");
            AssertCorpusFormula(sheet, 3, 2, -3d, "-A2");
            AssertCorpusFormula(sheet, 4, 2, 5d, "+A3");
            AssertCorpusFormula(sheet, 5, 2, 0.03d, "A2%");
            AssertCorpusFormula(sheet, 6, 2, true, "A1>A2");
            AssertCorpusFormula(sheet, 7, 2, true, "A2<=A3");
            AssertCorpusFormula(sheet, 8, 2, true, "A1<>A3");
            AssertCorpusFormula(sheet, 9, 2, 6d, "SUM({1,2,3})");
            AssertCorpusFormula(sheet, 10, 2, 3.33d, "ROUND(A1/A2,2)");
            AssertCorpusFormula(sheet, 11, 2, "yes", "IF(A1>A3,\"yes\",\"no\")");
            AssertCorpusFormula(sheet, 12, 2, 18d, "SUM(A1:A3)");
            AssertCorpusFormula(sheet, 1, 3, 103.33d, "B1+B10");
        }

        [Fact]
        public void LegacyXls_Corpus_Protection_PreservesSheetProtectionAndCellProtection() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "protection.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FILEPASS-UNSUPPORTED");
            Assert.Contains("PtgMul", result.ImportReport.FormulaTokensByName.Keys);

            Assert.NotNull(result.Workbook.Protection);
            Assert.True(result.Workbook.Protection!.IsProtected);
            Assert.Null(result.Workbook.Protection.LegacyPasswordHash);

            LegacyXlsWorksheet sheet = Assert.Single(result.Workbook.Worksheets);
            Assert.NotNull(sheet.Protection);
            Assert.True(sheet.Protection!.IsProtected);
            Assert.Matches("^[0-9A-F]{4}$", sheet.Protection.LegacyPasswordHash ?? string.Empty);
            AssertCorpusFormula(sheet, 2, 2, 14d, "A2*2");

            LegacyXlsCell inputCell = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 1);
            LegacyXlsCell formulaCell = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 2);
            LegacyXlsCellFormat inputFormat = Assert.Single(result.Workbook.CellFormats, format => format.StyleIndex == inputCell.StyleIndex);
            LegacyXlsCellFormat formulaFormat = Assert.Single(result.Workbook.CellFormats, format => format.StyleIndex == formulaCell.StyleIndex);
            Assert.True(inputFormat.ApplyProtection);
            Assert.False(inputFormat.Locked);
            Assert.False(inputFormat.FormulaHidden);
            Assert.True(formulaFormat.ApplyProtection);
            Assert.True(formulaFormat.Locked);
            Assert.True(formulaFormat.FormulaHidden);
        }

        [Fact]
        public void LegacyXls_DiagnosticCorpus_EncryptedWorkbook_ReportsFilePassBlocker() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsDiagnosticCorpus",
                "excel-com-generated",
                "encrypted-password.xls");

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Empty(workbook.Worksheets);
            Assert.Contains(workbook.Diagnostics, diagnostic =>
                diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error
                && diagnostic.Code == "XLS-BIFF-FILEPASS-UNSUPPORTED"
                && diagnostic.DetailCode == "Encryption:FilePass:Rc4");
            LegacyXlsUnsupportedFeature feature = Assert.Single(workbook.UnsupportedFeatures);
            Assert.Equal(LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook, feature.Kind);
            Assert.Equal("XLS-BIFF-FILEPASS-UNSUPPORTED", feature.Code);
            Assert.True(report.HasImportErrors);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook]);
            Assert.Equal(1, report.EncryptedWorkbooksByMethod["Rc4"]);
            Assert.Equal(1, report.FileFormatBlockers["EncryptedWorkbook|Encryption:FilePass:Rc4"]);
        }

        [Fact]
        public void LegacyXls_DiagnosticCorpus_Biff5Workbook_ReportsUnsupportedVersionBlocker() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsDiagnosticCorpus",
                "excel-com-generated",
                "biff5-workbook.xls");

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Empty(workbook.Worksheets);
            Assert.Contains(workbook.Diagnostics, diagnostic =>
                diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error
                && diagnostic.Code == "XLS-BIFF-VERSION-UNSUPPORTED"
                && diagnostic.DetailCode == "BiffVersion:BIFF5:WorkbookGlobals");
            LegacyXlsUnsupportedFeature feature = Assert.Single(workbook.UnsupportedFeatures);
            Assert.Equal(LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion, feature.Kind);
            Assert.Equal("XLS-BIFF-VERSION-UNSUPPORTED", feature.Code);
            Assert.True(report.HasImportErrors);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion]);
            Assert.Equal(1, report.UnsupportedBiffVersionsByVersion["BIFF5"]);
            Assert.Equal(1, report.UnsupportedBiffVersionsBySubstream["WorkbookGlobals"]);
            Assert.Equal(1, report.FileFormatBlockers["UnsupportedBiffVersion|BiffVersion:BIFF5:WorkbookGlobals"]);
        }

        private static bool IsLegacyXlsCorpusBaselineUpdateRequested() {
            string? value = Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_LEGACY_XLS_CORPUS_BASELINES");
            return string.Equals(value, "1", StringComparison.Ordinal)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static void AssertLegacyXlsCorpusBaselines(string corpusDirectory) {
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
                LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookPath, new LegacyXlsImportOptions {
                    ReportUnsupportedRecords = true
                });
                string actual = NormalizeBaselineText(workbook.CreateImportReport().ToMarkdown());
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
