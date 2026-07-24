using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
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

                Assert.True(document.SourceFormat == ExcelFileFormat.Xls);
                Assert.Equal(sourcePath, document.FilePath);
                Assert.DoesNotContain(document.LegacyXlsImportDiagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
                Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? header));
                Assert.Equal("Name", header);

                document.Save(outputPath);

                using ExcelDocument converted = ExcelDocument.Load(outputPath);
                Assert.False(converted.SourceFormat == ExcelFileFormat.Xls);
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

                Assert.True(document.SourceFormat == ExcelFileFormat.Xls);
                Assert.True(document.Sheets[0].TryGetCellText(2, 2, out string? amount));
                Assert.Equal("42", amount);
            } finally {
                TryDelete(sourcePath);
            }
        }

        [Fact]
        public void LegacyXls_NormalLoad_ReadOnlyModeUsesReadOnlyPackageAndRejectsSave() {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            using ExcelDocument document = ExcelDocument.Load(
                new MemoryStream(compound, writable: false),
                new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });

            Assert.Equal(OfficeIMO.Drawing.DocumentAccessMode.ReadOnly, document.AccessMode);
            Assert.Equal(FileAccess.Read, document.FileOpenAccess);
            Assert.Equal(ExcelFileFormat.Xls, document.SourceFormat);
            Assert.True(document.Sheets[0].TryGetCellText(2, 2, out string? value));
            Assert.Equal("42", value);

            using var destination = new MemoryStream();
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => document.Save(destination));
            Assert.Contains("read-only", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, destination.Length);
        }

        [Fact]
        public void LegacyXls_NormalLoad_LoadEncryptedPathDoesNotBindPlainSaveToEncryptedSource() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateRc4EncryptedWorkbookStream("openpass");
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string sourcePath = WriteTempWorkbook(compound, ".xls");

            try {
                using ExcelDocument document = ExcelDocument.LoadEncrypted(sourcePath, "openpass");

                Assert.True(document.SourceFormat == ExcelFileFormat.Xls);
                Assert.Null(document.FilePath);
                Assert.Equal("Rc4Sheet", document.Sheets.Single().Name);
                Assert.Throws<InvalidOperationException>(() => document.Save());
            } finally {
                TryDelete(sourcePath);
            }
        }

        [Fact]
        public async Task LegacyXls_NormalLoad_LoadEncryptedAsyncPathRoutesLegacyXls() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateRc4EncryptedWorkbookStream("openpass");
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string sourcePath = WriteTempWorkbook(compound, ".xls");

            try {
                using ExcelDocument document = await ExcelDocument.LoadEncryptedAsync(sourcePath, "openpass");

                Assert.True(document.SourceFormat == ExcelFileFormat.Xls);
                Assert.Null(document.FilePath);
                Assert.Equal("Rc4Sheet", document.Sheets.Single().Name);
                Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? value));
                Assert.Equal("RC4 secret", value);
            } finally {
                TryDelete(sourcePath);
            }
        }

        [Theory]
        [InlineData(true, false)]
        [InlineData(true, true)]
        [InlineData(false, false)]
        [InlineData(false, true)]
        public async Task EncryptedLoad_RoutesMisleadingExtensionsByContent(bool legacyPayload, bool asyncLoad) {
            const string password = "openpass";
            byte[] encryptedBytes;
            string misleadingExtension;

            if (legacyPayload) {
                byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateRc4EncryptedWorkbookStream(password);
                encryptedBytes = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
                misleadingExtension = ".xlsx";
            } else {
                using var sourceStream = new MemoryStream();
                using ExcelDocument source = ExcelDocument.Create(sourceStream);
                source.AddWorksheet("OpenXml").CellValue(1, 1, "Encrypted Open XML");
                using var encrypted = new MemoryStream();
                source.SaveEncrypted(encrypted, password);
                encryptedBytes = encrypted.ToArray();
                misleadingExtension = ".xls";
            }

            string sourcePath = WriteTempWorkbook(encryptedBytes, misleadingExtension);
            try {
                using ExcelDocument document = asyncLoad
                    ? await ExcelDocument.LoadEncryptedAsync(sourcePath, password)
                    : ExcelDocument.LoadEncrypted(sourcePath, password);

                Assert.Equal(legacyPayload ? ExcelFileFormat.Xls : ExcelFileFormat.Xlsx, document.SourceFormat);
                Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? value));
                Assert.Equal(legacyPayload ? "RC4 secret" : "Encrypted Open XML", value);
            } finally {
                TryDelete(sourcePath);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task EncryptedLoad_StreamRoutesLegacyXlsByContent(bool asyncLoad) {
            const string password = "openpass";
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateRc4EncryptedWorkbookStream(password);
            byte[] encryptedBytes = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            using var source = new MemoryStream(encryptedBytes);

            using ExcelDocument document = asyncLoad
                ? await ExcelDocument.LoadEncryptedAsync(source, password)
                : ExcelDocument.LoadEncrypted(source, password);

            Assert.Equal(ExcelFileFormat.Xls, document.SourceFormat);
            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? value));
            Assert.Equal("RC4 secret", value);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task EncryptedLoad_RejectsUnencryptedLegacyXls(bool asyncLoad) {
            byte[] compound = CreateMinimalLegacyXlsCompound();

            if (asyncLoad) {
                using var source = new MemoryStream(compound);
                InvalidDataException exception = await Assert.ThrowsAsync<InvalidDataException>(
                    () => ExcelDocument.LoadEncryptedAsync(source, "unused"));
                Assert.Contains("password-encrypted legacy XLS", exception.Message, StringComparison.Ordinal);
            } else {
                using var source = new MemoryStream(compound);
                InvalidDataException exception = Assert.Throws<InvalidDataException>(
                    () => ExcelDocument.LoadEncrypted(source, "unused"));
                Assert.Contains("password-encrypted legacy XLS", exception.Message, StringComparison.Ordinal);
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

            Assert.True(document.SourceFormat == ExcelFileFormat.Xls);
            Assert.NotNull(spreadsheet.WorkbookPart);
        }

        [Fact]
        public async Task LegacyXls_NormalLoad_SaveAsyncPathProducesValidXlsx() {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string sourcePath = WriteTempWorkbook(compound, ".xls");
            string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using ExcelDocument document = ExcelDocument.Load(sourcePath);

                await document.SaveAsync(outputPath);

                using ExcelDocument converted = ExcelDocument.Load(outputPath);
                Assert.False(converted.SourceFormat == ExcelFileFormat.Xls);
                Assert.True(converted.Sheets[0].TryGetCellText(1, 1, out string? header));
                Assert.Equal("Name", header);
                Assert.True(converted.Sheets[0].TryGetCellText(2, 2, out string? amount));
                Assert.Equal("42", amount);
            } finally {
                TryDelete(sourcePath);
                TryDelete(outputPath);
            }
        }

        [Fact]
        public async Task LegacyXls_NormalLoad_SaveAsyncStreamProducesValidXlsx() {
            byte[] compound = CreateMinimalLegacyXlsCompound();

            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(compound));
            using var output = new MemoryStream();

            await document.SaveAsync(output);

            output.Position = 0;
            using ExcelDocument converted = ExcelDocument.Load(output);
            Assert.False(converted.SourceFormat == ExcelFileFormat.Xls);
            Assert.True(converted.Sheets[0].TryGetCellText(1, 1, out string? header));
            Assert.Equal("Name", header);
            Assert.True(converted.Sheets[0].TryGetCellText(2, 2, out string? amount));
            Assert.Equal("42", amount);
        }

        [Fact]
        public void LegacyXls_NormalLoad_SaveEncryptedPathProducesEncryptedXlsx() {
            const string password = "legacy-xls-secret";
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string sourcePath = WriteTempWorkbook(compound, ".xls");
            string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using ExcelDocument document = ExcelDocument.Load(sourcePath);

                document.SaveEncrypted(outputPath, password);

                Assert.ThrowsAny<Exception>(() => SpreadsheetDocument.Open(outputPath, false).Dispose());
                using ExcelDocument decrypted = ExcelDocument.LoadEncrypted(outputPath, password);
                Assert.False(decrypted.SourceFormat == ExcelFileFormat.Xls);
                Assert.True(decrypted.Sheets[0].TryGetCellText(1, 1, out string? header));
                Assert.Equal("Name", header);
                Assert.True(decrypted.Sheets[0].TryGetCellText(2, 2, out string? amount));
                Assert.Equal("42", amount);
            } finally {
                TryDelete(sourcePath);
                TryDelete(outputPath);
            }
        }

        [Fact]
        public void LegacyXls_NormalLoad_SaveEncryptedStreamProducesEncryptedXlsx() {
            const string password = "legacy-xls-secret";
            byte[] compound = CreateMinimalLegacyXlsCompound();

            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(compound));
            using var output = new MemoryStream();

            document.SaveEncrypted(output, password);

            output.Position = 0;
            Assert.ThrowsAny<Exception>(() => SpreadsheetDocument.Open(output, false).Dispose());

            output.Position = 0;
            using ExcelDocument decrypted = ExcelDocument.LoadEncrypted(output, password);
            Assert.False(decrypted.SourceFormat == ExcelFileFormat.Xls);
            Assert.True(decrypted.Sheets[0].TryGetCellText(1, 1, out string? header));
            Assert.Equal("Name", header);
            Assert.True(decrypted.Sheets[0].TryGetCellText(2, 2, out string? amount));
            Assert.Equal("42", amount);
        }

        [Fact]
        public void LegacyXls_NormalLoad_RenamedOpenXmlWithXlsExtensionUsesOpenXmlLoader() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string renamedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("OpenXml");
                    sheet.CellValue(1, 1, "Open XML payload");
                    document.Save();
                }

                File.Copy(openXmlPath, renamedPath);

                using ExcelDocument renamed = ExcelDocument.Load(renamedPath);
                Assert.False(renamed.SourceFormat == ExcelFileFormat.Xls);
                Assert.Equal(renamedPath, renamed.FilePath);
                Assert.True(renamed.Sheets[0].TryGetCellText(1, 1, out string? value));
                Assert.Equal("Open XML payload", value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(renamedPath);
            }
        }

        [Fact]
        public void LegacyXls_NormalLoad_EmptyXlsExtensionThrowsClearCompoundDiagnostic() {
            string sourcePath = WriteTempWorkbook(Array.Empty<byte>(), ".xls");

            try {
                InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ExcelDocument.Load(sourcePath));

                Assert.Contains("Legacy XLS import failed", exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("XLS-COMPOUND-SIGNATURE", exception.Message, StringComparison.Ordinal);
            } finally {
                TryDelete(sourcePath);
            }
        }

        [Fact]
        public void LegacyXls_NormalLoad_ShortOleSignatureThrowsClearCompoundDiagnostic() {
            byte[] bytes = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ExcelDocument.Load(new MemoryStream(bytes)));

            Assert.Contains("Legacy XLS import failed", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("XLS-COMPOUND-SIGNATURE", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void LegacyXls_NormalLoad_CorruptOleCompoundThrowsClearCompoundDiagnostic() {
            byte[] bytes = LegacyXlsCompoundTestBuilder.CreateCompoundHeaderWithInvalidSectorChain();

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ExcelDocument.Load(new MemoryStream(bytes)));

            Assert.Contains("Legacy XLS import failed", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("XLS-COMPOUND-CORRUPT", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void LegacyXls_NormalLoad_NonExcelOleCompoundThrowsMissingWorkbookDiagnostic() {
            byte[] bytes = LegacyXlsCompoundTestBuilder.CreateNonExcelCompoundFile();

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ExcelDocument.Load(new MemoryStream(bytes)));

            Assert.Contains("Legacy XLS import failed", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("XLS-WORKBOOK-STREAM-MISSING", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task LegacyXls_NormalLoad_AsyncStreamProjectsToExcelDocument() {
            byte[] compound = CreateMinimalLegacyXlsCompound();

            using ExcelDocument document = await ExcelDocument.LoadAsync(new MemoryStream(compound));

            Assert.True(document.SourceFormat == ExcelFileFormat.Xls);
            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? header));
            Assert.Equal("Name", header);
        }

        [Fact]
        public void LegacyXls_NormalLoad_RejectsSaveOnDispose() {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string sourcePath = WriteTempWorkbook(compound, ".xls");

            try {
                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => ExcelDocument.Load(sourcePath, new OfficeIMO.Excel.ExcelLoadOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose }));

                Assert.Contains("SaveOnDispose is not supported", exception.Message, StringComparison.OrdinalIgnoreCase);
            } finally {
                TryDelete(sourcePath);
            }
        }

        [Fact]
        public void LegacyXls_NormalLoad_SavesNativeXlsTargets() {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string sourcePath = WriteTempWorkbook(compound, ".xls");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using ExcelDocument document = ExcelDocument.Load(sourcePath);

                document.Save();
                AssertNativeXlsRoundTrip(sourcePath);

                document.Save(xlsOutputPath);
                AssertNativeXlsRoundTrip(xlsOutputPath);
            } finally {
                TryDelete(sourcePath);
                TryDelete(xlsOutputPath);
            }
        }

        [Theory]
        [InlineData(".xlt")]
        [InlineData(".xla")]
        [InlineData(".xlm")]
        [InlineData(".xlw")]
        public void LegacyXls_NormalLoad_RejectsLegacyBinaryExcelSaveTargets(string extension) {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string sourcePath = WriteTempWorkbook(compound, extension);
            string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);

            try {
                using ExcelDocument document = ExcelDocument.Load(sourcePath);

                Assert.True(document.SourceFormat == ExcelFileFormat.Xls);
                Assert.Equal(sourcePath, document.FilePath);
                Assert.Throws<NotSupportedException>(() => document.Save());
                Assert.Throws<NotSupportedException>(() => document.Save(outputPath));
            } finally {
                TryDelete(sourcePath);
                TryDelete(outputPath);
            }
        }

        [Theory]
        [InlineData(".xlt")]
        [InlineData(".xla")]
        [InlineData(".xlm")]
        [InlineData(".xlw")]
        public void LegacyXls_ExplicitLoad_RejectsLegacyBinaryExcelSaveTargets(string extension) {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string sourcePath = WriteTempWorkbook(compound, extension);
            string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);

            try {
                using (ExcelDocument document = ExcelDocument.LoadLegacyXls(sourcePath)) {
                    Assert.True(document.SourceFormat == ExcelFileFormat.Xls);
                    Assert.Equal(sourcePath, document.FilePath);

                    AssertLegacyBinarySaveTargetRejected(() => document.Save());
                    AssertLegacyBinarySaveTargetRejected(() => document.Save(outputPath));
                }

                using (LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(sourcePath)) {
                    Assert.True(result.Document.SourceFormat == ExcelFileFormat.Xls);
                    Assert.Equal(sourcePath, result.Document.FilePath);

                    AssertLegacyBinarySaveTargetRejected(() => result.Document.Save());
                    AssertLegacyBinarySaveTargetRejected(() => result.Document.Save(outputPath));
                }
            } finally {
                TryDelete(sourcePath);
                TryDelete(outputPath);
            }
        }

        [Fact]
        public void LegacyXls_ExplicitLoad_SavesNativeXlsTargets() {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string sourcePath = WriteTempWorkbook(compound, ".xls");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using ExcelDocument document = ExcelDocument.LoadLegacyXls(sourcePath);

                Assert.True(document.SourceFormat == ExcelFileFormat.Xls);
                Assert.Equal(sourcePath, document.FilePath);
                document.Save();
                AssertNativeXlsRoundTrip(sourcePath);

                document.Save(xlsOutputPath);
                AssertNativeXlsRoundTrip(xlsOutputPath);
            } finally {
                TryDelete(sourcePath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_ExplicitLoadWithReport_SavesNativeXlsTargets() {
            byte[] compound = CreateMinimalLegacyXlsCompound();
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound));

                Assert.True(result.Document.SourceFormat == ExcelFileFormat.Xls);
                result.Document.Save(xlsOutputPath);
                AssertNativeXlsRoundTrip(xlsOutputPath);
            } finally {
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesBasicScalarWorkbook() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Data");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(2, 1, "Alice");
                    sheet.CellValue(2, 2, 42);
                    sheet.CellValue(3, 1, true);

                    document.Save(xlsOutputPath);
                }

                AssertNativeXlsRoundTrip(xlsOutputPath, expectedRow2Name: "Alice");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NormalAndExplicitLoad_ProjectChartOnlyWorkbooksAsXlsxChartSheets() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateChartOnlyWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using ExcelDocument normal = ExcelDocument.Load(new MemoryStream(compound));
                using ExcelDocument explicitLoad = ExcelDocument.LoadLegacyXls(new MemoryStream(compound));

                Assert.True(normal.SourceFormat == ExcelFileFormat.Xls);
                Assert.Empty(normal.LegacyXlsUnsupportedSheets);
                Assert.Empty(normal.LegacyXlsUnsupportedFeatures);
                Assert.Equal("ChartOnly", Assert.Single(normal.LegacyXlsChartSheets).Name);
                Assert.Equal("ChartOnly", Assert.Single(explicitLoad.LegacyXlsChartSheets).Name);

                normal.Save(outputPath);

                using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(outputPath, false);
                Assert.Single(spreadsheet.WorkbookPart!.ChartsheetParts);
                Assert.Equal(new[] { "ChartOnly" }, spreadsheet.WorkbookPart.Workbook.Sheets!.Elements<Sheet>().Select(sheet => sheet.Name?.Value).ToArray());
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));

                using ExcelDocumentReader reader = ExcelDocumentReader.Open(outputPath);
                Assert.Equal(0, reader.SheetCount);
                Assert.Empty(reader.GetSheetNames());
                Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetSheet(1));
                Assert.Throws<KeyNotFoundException>(() => reader.GetSheet("ChartOnly"));
            } finally {
                TryDelete(outputPath);
            }
        }

        [Fact]
        public void LegacyXls_NormalLoad_PreservesMixedWorksheetAndChartSheetOrder() {
            var workbook = new LegacyXlsWorkbook();
            workbook.MutableWorksheets.Add(new LegacyXlsWorksheet("DataBefore", streamOffset: 100, visibility: 0x00, sheetType: 0x00));
            workbook.MutableChartSheets.Add(new LegacyXlsChartSheet("ChartBetween", streamOffset: 200, visibility: 0x00, sheetType: 0x02));
            workbook.MutableWorksheets.Add(new LegacyXlsWorksheet("DataAfter", streamOffset: 300, visibility: 0x00, sheetType: 0x00));

            string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using (ExcelDocument document = workbook.ToExcelDocument()) {
                    document.Save(outputPath);
                }

                using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(outputPath, false);
                Assert.Equal(
                    new[] { "DataBefore", "ChartBetween", "DataAfter" },
                    spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Select(sheet => sheet.Name?.Value).ToArray());
                Assert.Single(spreadsheet.WorkbookPart.ChartsheetParts);
                Assert.Equal(2, spreadsheet.WorkbookPart.WorksheetParts.Count());
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
            } finally {
                TryDelete(outputPath);
            }
        }

        [Fact]
        public void LegacyXls_LoadLegacyXlsWithReport_ReturnsDocumentForChartOnlyWorkbooks() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateChartOnlyWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound));

            Assert.True(result.HasDocument);
            Assert.Null(result.ProjectionException);
            Assert.Equal(0, result.ImportReport.WorksheetCount);
            Assert.Equal(1, result.ImportReport.ChartSheetCount);
            Assert.Equal(0, result.ImportReport.UnsupportedSheetCount);
            Assert.False(result.HasImportErrors);
            Assert.Equal("ChartOnly", Assert.Single(result.Document.LegacyXlsChartSheets).Name);
        }

        [Fact]
        public void LegacyXls_LoadLegacyXlsWithReport_ReturnsDiagnosticsForHardImportErrors() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateEncryptedWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string sourcePath = WriteTempWorkbook(compound, ".xls");

            try {
                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(sourcePath);

                Assert.False(result.HasDocument);
                Assert.NotNull(result.ProjectionException);
                Assert.True(result.HasImportErrors);
                Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FILEPASS-UNSUPPORTED");
                Assert.Equal(0, result.ImportReport.WorksheetCount);

                InvalidOperationException documentException = Assert.Throws<InvalidOperationException>(() => result.Document);
                Assert.Contains("No OfficeIMO Excel document", documentException.Message, StringComparison.Ordinal);
                Assert.Same(result.ProjectionException, documentException.InnerException);
            } finally {
                TryDelete(sourcePath);
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

            Assert.True(document.SourceFormat == ExcelFileFormat.Xls);
            Assert.NotEmpty(document.LegacyXlsUnsupportedFeatures);

            ExcelFeatureReport report = document.InspectFeatures();
            ExcelFeatureFinding finding = Assert.Single(report.FindFeatures("Legacy XLS unsupported features"));

            Assert.Equal(ExcelFeatureSupportLevel.Unsupported, finding.SupportLevel);
            Assert.NotEmpty(finding.Details);
            Assert.True(report.HasAdvancedFeatures);
            Assert.Throws<InvalidOperationException>(() => report.EnsureNoUnsupportedFeatures());
            Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
        }

        [Fact]
        public void LegacyXls_LoadPolicyGuards_ShowHowCallersCanRejectUnsafeConversions() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            result.EnsureNoImportErrors();
            InvalidOperationException unsupported = Assert.Throws<InvalidOperationException>(() => result.EnsureNoUnsupportedFeatures());
            Assert.Contains("unsupported or preserve-only features", unsupported.Message, StringComparison.OrdinalIgnoreCase);

            ExcelFeatureReport featureReport = result.Document.InspectFeatures();
            Assert.Throws<InvalidOperationException>(() => featureReport.EnsureNoUnsupportedFeatures());
            Assert.Throws<InvalidOperationException>(() => featureReport.EnsureNoAdvancedFeatures());
        }

        private static void AssertLegacyBinarySaveTargetRejected(Action saveAction) {
            NotSupportedException exception = Assert.Throws<NotSupportedException>(saveAction);
            Assert.Contains(".xls workbook files only", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(".xlt, .xla, .xlm, and .xlw", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        private static byte[] CreateMinimalLegacyXlsCompound() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            return LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
        }

        private static void AssertNativeXlsRoundTrip(string path, string expectedName = "Name", string expectedAmount = "42", string expectedFlag = "1", string? expectedRow2Name = null) {
            byte[] bytes = File.ReadAllBytes(path);
            Assert.True(bytes.Length > 8);
            Assert.Equal(0xd0, bytes[0]);
            Assert.Equal(0xcf, bytes[1]);
            Assert.Equal(0x11, bytes[2]);
            Assert.Equal(0xe0, bytes[3]);

            using ExcelDocument normal = ExcelDocument.Load(path);
            Assert.True(normal.SourceFormat == ExcelFileFormat.Xls);
            Assert.True(normal.Sheets[0].TryGetCellText(1, 1, out string? name));
            Assert.Equal(expectedName, name);
            if (expectedRow2Name != null) {
                Assert.True(normal.Sheets[0].TryGetCellText(2, 1, out string? row2Name));
                Assert.Equal(expectedRow2Name, row2Name);
            }

            Assert.True(normal.Sheets[0].TryGetCellText(2, 2, out string? amount));
            Assert.Equal(expectedAmount, amount);
            Assert.True(normal.Sheets[0].TryGetCellText(3, 1, out string? flag));
            Assert.Equal(expectedFlag, flag);

            using ExcelDocument explicitLoad = ExcelDocument.LoadLegacyXls(path);
            Assert.True(explicitLoad.SourceFormat == ExcelFileFormat.Xls);
            Assert.True(explicitLoad.Sheets[0].TryGetCellText(1, 1, out string? explicitName));
            Assert.Equal(expectedName, explicitName);
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

        private static partial class LegacyXlsTestWorkbookBuilder {
            internal static byte[] CreateChartOnlyWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long chartBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ChartOnly", sheetType: 0x02));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int chartSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x20, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(chartSheetOffset), 0, bytes, checked((int)chartBoundSheetPosition + 4), 4);
                return bytes;
            }
        }
    }
}
