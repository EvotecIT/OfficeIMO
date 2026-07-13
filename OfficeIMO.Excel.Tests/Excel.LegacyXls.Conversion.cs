using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_Convert_XlsxToXlsAndBack_RoundTripsSupportedContent() {
            string xlsxPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xls");
            string roundTripPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");

            using (ExcelDocument document = ExcelDocument.Create(xlsxPath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alice");
                sheet.CellValue(2, 2, 42);
                sheet.CellValue(3, 1, true);
                document.Save();
            }

            ExcelDocumentConversionResult toXls = ExcelDocument.Convert(xlsxPath, xlsPath);

            Assert.Equal(xlsPath, toXls.RequireNoLoss());
            Assert.Equal(ExcelFileFormat.Xlsx, toXls.Report.SourceFormat);
            Assert.Equal(ExcelFileFormat.Xls, toXls.Report.DestinationFormat);
            Assert.False(toXls.HasLoss);

            AssertNativeXlsRoundTrip(xlsPath, expectedRow2Name: "Alice");

            ExcelDocumentConversionResult toXlsx = ExcelDocument.Convert(xlsPath, roundTripPath);

            Assert.Equal(ExcelFileFormat.Xls, toXlsx.Report.SourceFormat);
            Assert.Equal(ExcelFileFormat.Xlsx, toXlsx.Report.DestinationFormat);

            using ExcelDocument roundTrip = ExcelDocument.Load(roundTripPath);
            Assert.False(roundTrip.SourceFormat == ExcelFileFormat.Xls);
            Assert.True(roundTrip.Sheets[0].TryGetCellText(1, 1, out string? header));
            Assert.Equal("Name", header);
            Assert.True(roundTrip.Sheets[0].TryGetCellText(2, 1, out string? name));
            Assert.Equal("Alice", name);
            Assert.True(roundTrip.Sheets[0].TryGetCellText(2, 2, out string? amount));
            Assert.Equal("42", amount);
        }

        [Fact]
        public void LegacyXls_Convert_BlocksUnsupportedLegacyContentUnlessLossIsAllowed() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string xlsPath = WriteTempWorkbook(compound, ".xls");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                ExcelDocumentConversionException exception = Assert.Throws<ExcelDocumentConversionException>(() => ExcelDocument.Convert(xlsPath, blockedPath));

                Assert.Equal(ExcelDocumentConversionFailureReason.DataLossBlocked, exception.Reason);
                Assert.True(exception.Result.HasLoss);
                Assert.False(File.Exists(blockedPath));

                ExcelDocumentConversionResult result = ExcelDocument.Convert(xlsPath, allowedPath, new ExcelDocumentConversionOptions {
                    LossPolicy = ExcelConversionLossPolicy.Allow
                });

                Assert.True(result.HasLoss);
                Assert.Equal(allowedPath, result.RequireValue());
                Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());

                using ExcelDocument converted = ExcelDocument.Load(allowedPath);
                Assert.False(converted.SourceFormat == ExcelFileFormat.Xls);
                Assert.True(converted.Sheets[0].TryGetCellText(1, 1, out string? header));
                Assert.Equal("Feature", header);
            } finally {
                TryDelete(xlsPath);
            }
        }

        [Fact]
        public void LegacyXls_Convert_ContentDetectionAppliesImportOptionsDespiteMisleadingExtension() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string sourcePath = WriteTempWorkbook(compound, ".xlsx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                    ExcelDocument.Convert(sourcePath, destinationPath, new ExcelDocumentConversionOptions {
                        LossPolicy = ExcelConversionLossPolicy.Allow,
                        LegacyXlsImportOptions = new LegacyXlsImportOptions { MaxInputBytes = 1 }
                    }));

                Assert.Contains("configured limit of 1 byte", exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.False(File.Exists(destinationPath));
            } finally {
                TryDelete(sourcePath);
            }
        }

        [Fact]
        public void LegacyXls_Convert_CannotSuppressDiscoveryToBypassLossPolicy() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string sourcePath = WriteTempWorkbook(compound, ".xls");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                ExcelDocumentConversionException exception = Assert.Throws<ExcelDocumentConversionException>(() =>
                    ExcelDocument.Convert(sourcePath, destinationPath, new ExcelDocumentConversionOptions {
                        LegacyXlsImportOptions = new LegacyXlsImportOptions { ReportUnsupportedContent = false }
                    }));

                Assert.Equal(ExcelDocumentConversionFailureReason.DataLossBlocked, exception.Reason);
                Assert.True(exception.Result.HasLoss);
                Assert.False(File.Exists(destinationPath));
            } finally {
                TryDelete(sourcePath);
            }
        }

        [Fact]
        public void LegacyXls_Convert_DefaultConflictPolicyPreservesExistingDestination() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xls");
            byte[] existing = { 1, 2, 3, 4 };
            using (ExcelDocument document = ExcelDocument.Create(sourcePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Conflict policy");
                document.Save();
            }
            File.WriteAllBytes(destinationPath, existing);

            ExcelDocumentConversionException exception = Assert.Throws<ExcelDocumentConversionException>(() =>
                ExcelDocument.Convert(sourcePath, destinationPath));

            Assert.Equal(ExcelDocumentConversionFailureReason.DestinationExists, exception.Reason);
            Assert.Equal(existing, File.ReadAllBytes(destinationPath));

            ExcelDocumentConversionResult replaced = ExcelDocument.Convert(sourcePath, destinationPath, new ExcelDocumentConversionOptions {
                FileConflictPolicy = ExcelConversionFileConflictPolicy.Replace
            });
            Assert.True(replaced.Report.ReplacedExistingFile);
            using ExcelDocument loaded = ExcelDocument.Load(destinationPath);
            Assert.True(loaded.SourceFormat == ExcelFileFormat.Xls);
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string? value));
            Assert.Equal("Conflict policy", value);
        }

        [Fact]
        public void LegacyXls_Convert_ReplacePreservesReadOnlyDestination() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xls");
            byte[] originalBytes = { 1, 2, 3, 4 };
            using (ExcelDocument document = ExcelDocument.Create(sourcePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Read-only conversion target");
                document.Save();
            }
            File.WriteAllBytes(destinationPath, originalBytes);
            var destination = new FileInfo(destinationPath) { IsReadOnly = true };

            try {
                Assert.Throws<IOException>(() => ExcelDocument.Convert(sourcePath, destinationPath, new ExcelDocumentConversionOptions {
                    FileConflictPolicy = ExcelConversionFileConflictPolicy.Replace
                }));

                Assert.Equal(originalBytes, File.ReadAllBytes(destinationPath));
            } finally {
                destination.IsReadOnly = false;
            }
        }

        [Fact]
        public void LegacyXls_Convert_RejectsSamePhysicalFormat() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(sourcePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Already XLSX");
                document.Save();
            }

            ExcelDocumentConversionException exception = Assert.Throws<ExcelDocumentConversionException>(() =>
                ExcelDocument.Convert(sourcePath, destinationPath));

            Assert.Equal(ExcelDocumentConversionFailureReason.SameFormat, exception.Reason);
            Assert.False(File.Exists(destinationPath));
        }

        [Fact]
        public void LegacyXls_Convert_DisablesOpenSettingsAutoSaveForSourceIsolation() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(sourcePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Source must remain untouched");
                document.Save();
            }
            File.SetLastWriteTimeUtc(sourcePath, new DateTime(2001, 1, 1, 0, 0, 0, DateTimeKind.Utc));
            byte[] sourceBytes = File.ReadAllBytes(sourcePath);
            DateTime sourceWriteTime = File.GetLastWriteTimeUtc(sourcePath);

            ExcelDocumentConversionException exception = Assert.Throws<ExcelDocumentConversionException>(() =>
                ExcelDocument.Convert(sourcePath, destinationPath, new ExcelDocumentConversionOptions {
                    OpenSettings = new OpenSettings { AutoSave = true }
                }));

            Assert.Equal(ExcelDocumentConversionFailureReason.SameFormat, exception.Reason);
            Assert.Equal(sourceBytes, File.ReadAllBytes(sourcePath));
            Assert.Equal(sourceWriteTime, File.GetLastWriteTimeUtc(sourcePath));
            Assert.False(File.Exists(destinationPath));
        }

        [Fact]
        public void LegacyXls_Convert_BlocksCompoundFeatureMetadataUnlessLossIsAllowed() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithOleObjectStorage(workbookStream);
            string xlsPath = WriteTempWorkbook(compound, ".xls");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                ExcelDocumentConversionException exception = Assert.Throws<ExcelDocumentConversionException>(() => ExcelDocument.Convert(xlsPath, blockedPath));

                Assert.Equal(ExcelDocumentConversionFailureReason.DataLossBlocked, exception.Reason);
                Assert.Contains(exception.Result.Report.Diagnostics, diagnostic => diagnostic.Code.Contains("Compound", StringComparison.Ordinal));
                Assert.False(File.Exists(blockedPath));

                ExcelDocument.Convert(xlsPath, allowedPath, new ExcelDocumentConversionOptions {
                    LossPolicy = ExcelConversionLossPolicy.Allow
                });

                using ExcelDocument converted = ExcelDocument.Load(allowedPath);
                Assert.False(converted.SourceFormat == ExcelFileFormat.Xls);
                Assert.True(converted.Sheets[0].TryGetCellText(1, 1, out string? text));
                Assert.Equal("Name", text);
            } finally {
                TryDelete(xlsPath);
            }
        }

        [Fact]
        public void LegacyXls_Convert_ProjectsChartSheetsWithoutLossyOverride() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateChartOnlyWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string xlsPath = WriteTempWorkbook(compound, ".xls");
            string convertedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                ExcelDocument.Convert(xlsPath, convertedPath);

                using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(convertedPath, false);
                Assert.Single(spreadsheet.WorkbookPart!.ChartsheetParts);
                Assert.Contains(spreadsheet.WorkbookPart.Workbook.Sheets!.Elements<Sheet>(), sheet => sheet.Name?.Value == "ChartOnly");
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
            } finally {
                TryDelete(xlsPath);
                TryDelete(convertedPath);
            }
        }

        [Fact]
        public void LegacyXls_NormalSave_BlocksKnownImportLossUnlessExplicitlyAllowed() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string xlsPath = WriteTempWorkbook(compound, ".xls");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using ExcelDocument document = ExcelDocument.Load(xlsPath);

                Assert.Throws<NotSupportedException>(() => document.Save(blockedPath));
                Assert.False(File.Exists(blockedPath));

                document.Save(allowedPath, new ExcelSaveOptions {
                    LossPolicy = ExcelConversionLossPolicy.Allow
                });

                using ExcelDocument saved = ExcelDocument.Load(allowedPath);
                Assert.True(saved.Sheets[0].TryGetCellText(1, 1, out string? text));
                Assert.Equal("Feature", text);
            } finally {
                TryDelete(xlsPath);
            }
        }

        [Fact]
        public void LegacyXls_NormalSave_BlocksPreservedRecordsWithoutUnsupportedSummary() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string xlsPath = WriteTempWorkbook(compound, ".xls");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using ExcelDocument document = ExcelDocument.Load(xlsPath);
                Assert.NotEmpty(document.LegacyXlsPreservedFeatures);

                typeof(ExcelDocument)
                    .GetField("_legacyXlsUnsupportedFeatures", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
                    .SetValue(document, Array.Empty<LegacyXlsUnsupportedFeature>());

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(destinationPath));

                Assert.Contains(nameof(ExcelDocument.LegacyXlsPreservedFeatures), exception.Message, StringComparison.Ordinal);
                Assert.False(File.Exists(destinationPath));
            } finally {
                TryDelete(xlsPath);
                TryDelete(destinationPath);
            }
        }

        [Fact]
        public void LegacyXls_NormalSave_DoesNotTreatProjectedChartSheetsAsLoss() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateChartOnlyWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string xlsPath = WriteTempWorkbook(compound, ".xls");
            string xlsxPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using (ExcelDocument document = ExcelDocument.Load(xlsPath)) {
                    document.Save(xlsxPath);
                }

                using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(xlsxPath, false);
                Assert.Single(spreadsheet.WorkbookPart!.ChartsheetParts);
            } finally {
                TryDelete(xlsPath);
                TryDelete(xlsxPath);
            }
        }
    }
}
