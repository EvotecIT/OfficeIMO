using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NormalLoad_PathProjectsToExcelDocumentAndSavesXlsx() {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string sourcePath = WriteTempWorkbook(compound, ".xls");
            string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using ExcelDocument document = ExcelDocument.Load(sourcePath);

                Assert.True(document.WasLoadedFromLegacyXls);
                Assert.Equal(sourcePath, document.FilePath);
                Assert.DoesNotContain(document.LegacyXlsImportDiagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
                Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? header));
                Assert.Equal("Name", header);

                document.Save(outputPath);

                using ExcelDocument converted = ExcelDocument.Load(outputPath);
                Assert.False(converted.WasLoadedFromLegacyXls);
                Assert.True(converted.Sheets[0].TryGetCellText(2, 2, out string? amount));
                Assert.Equal("42", amount);
            } finally {
                TryDelete(sourcePath);
                TryDelete(outputPath);
            }
        }

        [Fact]
        public async Task LegacyXls_NormalLoad_AsyncPathProjectsToExcelDocument() {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string sourcePath = WriteTempWorkbook(compound, ".xls");

            try {
                using ExcelDocument document = await ExcelDocument.LoadAsync(sourcePath);

                Assert.True(document.WasLoadedFromLegacyXls);
                Assert.True(document.Sheets[0].TryGetCellText(2, 2, out string? amount));
                Assert.Equal("42", amount);
            } finally {
                TryDelete(sourcePath);
            }
        }

        [Fact]
        public void LegacyXls_NormalLoad_StreamProjectsToExcelDocumentAndSavesOpenXmlStream() {
            byte[] compound = CreateMinimalLegacyXlsCompound();

            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(compound));
            using var output = new MemoryStream();

            document.Save(output);
            output.Position = 0;
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(output, false);

            Assert.True(document.WasLoadedFromLegacyXls);
            Assert.NotNull(spreadsheet.WorkbookPart);
        }

        [Fact]
        public async Task LegacyXls_NormalLoad_AsyncStreamProjectsToExcelDocument() {
            byte[] compound = CreateMinimalLegacyXlsCompound();

            using ExcelDocument document = await ExcelDocument.LoadAsync(new MemoryStream(compound));

            Assert.True(document.WasLoadedFromLegacyXls);
            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? header));
            Assert.Equal("Name", header);
        }

        [Fact]
        public void LegacyXls_NormalLoad_RejectsAutoSave() {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string sourcePath = WriteTempWorkbook(compound, ".xls");

            try {
                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => ExcelDocument.Load(sourcePath, autoSave: true));

                Assert.Contains("Auto-save is not supported", exception.Message, StringComparison.OrdinalIgnoreCase);
            } finally {
                TryDelete(sourcePath);
            }
        }

        [Fact]
        public void LegacyXls_NormalLoad_RejectsNativeXlsSaveTargets() {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string sourcePath = WriteTempWorkbook(compound, ".xls");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using ExcelDocument document = ExcelDocument.Load(sourcePath);

                NotSupportedException implicitSave = Assert.Throws<NotSupportedException>(() => document.Save());
                NotSupportedException explicitSave = Assert.Throws<NotSupportedException>(() => document.Save(xlsOutputPath));

                Assert.Contains("Native XLS saving is not supported", implicitSave.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("Native XLS saving is not supported", explicitSave.Message, StringComparison.OrdinalIgnoreCase);
            } finally {
                TryDelete(sourcePath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_ExplicitLoad_RejectsNativeXlsSaveTargets() {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string sourcePath = WriteTempWorkbook(compound, ".xls");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using ExcelDocument document = ExcelDocument.LoadLegacyXls(sourcePath);

                Assert.True(document.WasLoadedFromLegacyXls);
                Assert.Equal(sourcePath, document.FilePath);
                Assert.Throws<NotSupportedException>(() => document.Save());
                Assert.Throws<NotSupportedException>(() => document.Save(xlsOutputPath));
            } finally {
                TryDelete(sourcePath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_ExplicitLoadWithReport_RejectsNativeXlsSaveTargets() {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound));

                Assert.True(result.Document.WasLoadedFromLegacyXls);
                Assert.Throws<NotSupportedException>(() => result.Document.Save(xlsOutputPath));
            } finally {
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NormalLoad_ThrowsForHardImportErrors() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateEncryptedWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string sourcePath = WriteTempWorkbook(compound, ".xls");

            try {
                InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ExcelDocument.Load(sourcePath));

                Assert.Contains("Legacy XLS import failed", exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("XLS-BIFF-FILEPASS-UNSUPPORTED", exception.Message, StringComparison.Ordinal);
            } finally {
                TryDelete(sourcePath);
            }
        }

        [Fact]
        public void LegacyXls_NormalLoad_ExposesImportDiagnosticsThroughFeatureReport() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(compound));

            Assert.True(document.WasLoadedFromLegacyXls);
            Assert.NotEmpty(document.LegacyXlsUnsupportedFeatures);

            ExcelFeatureReport report = document.InspectFeatures();
            ExcelFeatureFinding finding = Assert.Single(report.FindFeatures("Legacy XLS preserve-only features"));

            Assert.Equal(ExcelFeatureSupportLevel.Preserved, finding.SupportLevel);
            Assert.NotEmpty(finding.Details);
            Assert.True(report.HasAdvancedFeatures);
            Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
        }

        private static byte[] CreateMinimalLegacyXlsCompound() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            return LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
        }

        private static string WriteTempWorkbook(byte[] bytes, string extension) {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
            File.WriteAllBytes(path, bytes);
            return path;
        }

        private static void TryDelete(string path) {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }
}
