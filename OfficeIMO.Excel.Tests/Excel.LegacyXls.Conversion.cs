using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
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
        public void Convert_ModernWorkbookToTemplate_UsesConcreteFormatDescriptors() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xltx");
            using (ExcelDocument document = ExcelDocument.Create(sourcePath)) {
                document.AddWorksheet("Template").CellValue(1, 1, "Reusable workbook");
                document.Save();
            }

            ExcelDocumentConversionResult result = ExcelDocument.Convert(sourcePath, destinationPath);

            Assert.Equal("Excel.Xlsx", result.Report.SourceFormatDescriptor.Id);
            Assert.Equal("Excel.Xltx", result.Report.DestinationFormatDescriptor.Id);
            Assert.Equal(OfficeDocumentKind.Template, result.Report.DestinationFormatDescriptor.DocumentKind);
            Assert.Equal(destinationPath, result.RequireNoLoss());
            using SpreadsheetDocument package = SpreadsheetDocument.Open(destinationPath, false);
            Assert.Equal(SpreadsheetDocumentType.Template, package.DocumentType);
        }

        [Fact]
        public void Convert_MacroWorkbookToMacroFreeWorkbook_BlocksOrReportsRemoval() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsm");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(sourcePath)) {
                document.AddWorksheet("Macro").CellValue(1, 1, "Carrier");
                document.AddMacro(CreateVbaProjectPayload());
                document.Save();
            }

            ExcelDocumentConversionException blocked = Assert.Throws<ExcelDocumentConversionException>(() =>
                ExcelDocument.Convert(sourcePath, blockedPath));

            Assert.Equal(ExcelDocumentConversionFailureReason.DataLossBlocked, blocked.Reason);
            ExcelConversionDiagnostic finding = Assert.Single(blocked.Result.Report.Diagnostics,
                diagnostic => diagnostic.Code == "Excel.VbaProject.Removed");
            Assert.Equal(OfficeCompatibilityState.Blocked, finding.CompatibilityState);
            Assert.True(finding.CompatibilityImpact.HasFlag(OfficeCompatibilityImpact.Security));
            Assert.False(File.Exists(blockedPath));

            ExcelDocumentConversionResult allowed = ExcelDocument.Convert(sourcePath, allowedPath,
                new ExcelDocumentConversionOptions { CompatibilityMode = OfficeCompatibilityMode.BestEffort });

            Assert.True(allowed.HasLoss);
            Assert.Equal(OfficeCompatibilityMode.BestEffort, allowed.Report.Compatibility.Mode);
            using ExcelDocument converted = ExcelDocument.Load(allowedPath);
            Assert.False(converted.HasMacros);
        }

        [Fact]
        public void Convert_ClassifiedLegacyTemplateDestination_IsExplicitlyBlocked() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlt");
            using (ExcelDocument document = ExcelDocument.Create(sourcePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "No silent extension aliasing");
                document.Save();
            }

            ExcelDocumentConversionException exception = Assert.Throws<ExcelDocumentConversionException>(() =>
                ExcelDocument.Convert(sourcePath, destinationPath));

            Assert.Equal(ExcelDocumentConversionFailureReason.DestinationFeatureUnsupported, exception.Reason);
            Assert.Contains(exception.Result.Report.Diagnostics,
                diagnostic => diagnostic.Code == "Excel.LegacyDestination.NotWritable"
                    && diagnostic.CompatibilityState == OfficeCompatibilityState.Blocked);
            Assert.False(File.Exists(destinationPath));
        }

        [Fact]
        public void AnalyzeConversion_ReportsBlockedTargetWithoutCreatingOutput() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlt");
            using (ExcelDocument document = ExcelDocument.Create(sourcePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Preflight only");
                document.Save();
            }

            var options = new ExcelDocumentConversionOptions {
                CompatibilityMode = OfficeCompatibilityMode.PreservationOnly
            };
            ExcelDocumentConversionReport report = ExcelDocument.AnalyzeConversion(sourcePath, destinationPath, options);

            Assert.True(report.Compatibility.HasBlockedFeatures);
            Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "Excel.LegacyDestination.NotWritable");
            Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == "Excel.SourceCarrier.Embedded");
            Assert.False(File.Exists(destinationPath));

            ExcelDocumentConversionException exception = Assert.Throws<ExcelDocumentConversionException>(() =>
                ExcelDocument.Convert(sourcePath, destinationPath, options));
            Assert.DoesNotContain(exception.Result.Report.Diagnostics,
                diagnostic => diagnostic.Code == "Excel.SourceCarrier.Embedded");
            Assert.False(File.Exists(destinationPath));
        }

        [Fact]
        public void Convert_SignedModernSourceIsIncludedInAnalysisAndStructuredFailure() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xltx");
            string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xltx");
            using (ExcelDocument document = ExcelDocument.Create(sourcePath)) {
                document.AddWorksheet("Signed").CellValue(1, 1, "Signed conversion source");
                document.Save();
            }
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(sourcePath, true)) {
                spreadsheet.AddDigitalSignatureOriginPart();
                DigitalSignatureOriginPart originPart = spreadsheet.DigitalSignatureOriginPart!;
                XmlSignaturePart signaturePart = originPart.AddNewPart<XmlSignaturePart>();
                using var signatureStream = new MemoryStream(Encoding.UTF8.GetBytes(
                    "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo /></Signature>"));
                signaturePart.FeedData(signatureStream);
                ExtendedFilePropertiesPart appPart = spreadsheet.ExtendedFilePropertiesPart
                    ?? spreadsheet.AddExtendedFilePropertiesPart();
                appPart.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
                appPart.Properties.DigitalSignature = new DocumentFormat.OpenXml.ExtendedProperties.DigitalSignature();
                appPart.Properties.Save();
            }

            ExcelDocumentConversionReport analysis = ExcelDocument.AnalyzeConversion(sourcePath, blockedPath);
            Assert.True(analysis.Compatibility.HasSecurityImpact);
            Assert.Contains(analysis.Compatibility.Findings,
                finding => finding.Code == "Excel.DigitalSignature.Invalidated"
                    && finding.State == OfficeCompatibilityState.Blocked
                    && finding.RepresentsLoss);

            ExcelDocumentConversionException blocked = Assert.Throws<ExcelDocumentConversionException>(() =>
                ExcelDocument.Convert(sourcePath, blockedPath));
            Assert.Equal(ExcelDocumentConversionFailureReason.DataLossBlocked, blocked.Reason);
            Assert.False(File.Exists(blockedPath));

            ExcelDocumentConversionResult allowed = ExcelDocument.Convert(
                sourcePath,
                allowedPath,
                new ExcelDocumentConversionOptions {
                    CompatibilityMode = OfficeCompatibilityMode.BestEffort,
                    SaveOptions = new ExcelSaveOptions {
                        SignatureMutationPolicy = ExcelSignatureMutationPolicy.RemoveInvalidatedSignatures
                    }
                });
            Assert.True(allowed.Report.Compatibility.HasSecurityImpact);
            Assert.Contains(allowed.Report.Compatibility.Findings,
                finding => finding.Code == "Excel.DigitalSignature.Invalidated"
                    && finding.State == OfficeCompatibilityState.Dropped
                    && finding.RepresentsLoss);
            using SpreadsheetDocument converted = SpreadsheetDocument.Open(allowedPath, false);
            Assert.Null(converted.DigitalSignatureOriginPart);
            Assert.Null(converted.ExtendedFilePropertiesPart?.Properties?.DigitalSignature);
        }

        [Theory]
        [InlineData(".xlt")]
        [InlineData(".xla")]
        [InlineData(".xlm")]
        [InlineData(".xlw")]
        public void SaveApis_RejectClassifiedButUnwritableLegacyVariantsBeforeWriting(string extension) {
            string savePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + extension);
            string copyPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + extension);
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data").CellValue(1, 1, "No mislabeled Open XML payload");

            NotSupportedException save = Assert.Throws<NotSupportedException>(() => document.Save(savePath));
            NotSupportedException copy = Assert.Throws<NotSupportedException>(() => document.SaveCopy(copyPath));

            Assert.Contains("not supported", save.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("supports", copy.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(savePath));
            Assert.False(File.Exists(copyPath));
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

        [Fact]
        public void BinaryConversion_ChartUsesPolicyDrivenCellRasterFallback() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xls");
            string visualPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xls");
            using (ExcelDocument document = ExcelDocument.Create(sourcePath)) {
                ExcelSheet sheet = document.AddWorksheet("Dashboard");
                sheet.CellValue(1, 1, "Quarter");
                sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(2, 1, "Q1");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(3, 1, "Q2");
                sheet.CellValue(3, 2, 20);
                sheet.AddChart(
                    new ExcelChartData(
                        new[] { "Q1", "Q2" },
                        new[] { new ExcelChartSeries("Sales", new[] { 10D, 20D }) }),
                    row: 1,
                    column: 4,
                    widthPixels: 280,
                    heightPixels: 180,
                    type: ExcelChartType.ColumnClustered,
                    title: "Quarterly sales");
                document.Save();
            }

            ExcelDocumentConversionException blocked = Assert.Throws<ExcelDocumentConversionException>(() =>
                ExcelDocument.Convert(sourcePath, blockedPath));

            Assert.Equal(ExcelDocumentConversionFailureReason.DestinationFeatureUnsupported, blocked.Reason);
            Assert.Contains(blocked.Result.Report.Diagnostics,
                finding => finding.Code == "Excel.BinaryWriter.Unsupported"
                    && finding.CompatibilityState == OfficeCompatibilityState.Blocked);
            Assert.False(File.Exists(blockedPath));

            var options = new ExcelDocumentConversionOptions {
                CompatibilityMode = OfficeCompatibilityMode.PreferVisual,
                VisualFallbackMaxColumns = 32,
                VisualFallbackMaxRows = 64
            };
            ExcelDocumentConversionReport preflight = ExcelDocument.AnalyzeConversion(sourcePath, visualPath, options);

            Assert.False(File.Exists(visualPath));
            Assert.Contains(preflight.Diagnostics,
                finding => finding.Code == "Excel.BinaryWriter.CellRasterFallback"
                    && finding.CompatibilityState == OfficeCompatibilityState.Rasterized
                    && finding.FallbackArtifact != null);
            ExcelDocumentConversionResult converted = ExcelDocument.Convert(sourcePath, visualPath, options);

            Assert.Equal(
                preflight.Diagnostics.Select(finding => finding.Code),
                converted.Report.Diagnostics.Select(finding => finding.Code));
            using ExcelDocument legacy = ExcelDocument.Load(visualPath);
            Assert.Equal(ExcelFileFormat.Xls, legacy.SourceFormat);
            ExcelSheet visualSheet = Assert.Single(legacy.Sheets, sheet => sheet.Name == "Dashboard");
            Assert.False(string.IsNullOrWhiteSpace(visualSheet.CellAt(1, 1).GetStyle().FillColorArgb));
            Assert.False(legacy.TryGetCompatibilitySourcePayload(out _, out _));
        }

        [Fact]
        public void BinaryConversion_PreservationOnlyRetainsOriginalInXlsbCarrier() {
            string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsb");
            using (ExcelDocument document = ExcelDocument.Create(sourcePath)) {
                ExcelSheet sheet = document.AddWorksheet("Visual");
                sheet.CellValue(1, 1, "Source payload");
                sheet.AddChart(
                    new ExcelChartData(
                        new[] { "A", "B" },
                        new[] { new ExcelChartSeries("Value", new[] { 3D, 7D }) }),
                    row: 2,
                    column: 2,
                    widthPixels: 220,
                    heightPixels: 140,
                    type: ExcelChartType.Pie,
                    title: "Preserved chart");
                document.Save();
            }
            byte[] sourceBytes = File.ReadAllBytes(sourcePath);

            ExcelDocumentConversionResult converted = ExcelDocument.Convert(
                sourcePath,
                destinationPath,
                new ExcelDocumentConversionOptions {
                    CompatibilityMode = OfficeCompatibilityMode.PreservationOnly,
                    VisualFallbackMaxColumns = 24,
                    VisualFallbackMaxRows = 48
                });

            Assert.Contains(converted.Report.Diagnostics,
                finding => finding.Code == "Excel.SourceCarrier.Embedded"
                    && finding.CompatibilityState == OfficeCompatibilityState.EmbeddedSource);
            using ExcelDocument binary = ExcelDocument.Load(destinationPath);
            Assert.Equal(ExcelFileFormat.Xlsb, binary.SourceFormat);
            Assert.True(binary.TryGetCompatibilitySourcePayload(out OfficeCompatibilitySourcePayload? payload, out string? error), error);
            Assert.NotNull(payload);
            Assert.Equal("Excel.Xlsx", payload!.FormatId);
            Assert.Equal(OfficeCompatibilityMode.PreservationOnly, payload.Mode);
            Assert.Equal(sourceBytes, payload.ToArray());
        }

        [Fact]
        public void LegacyXls_Convert_PreservationOnlyRetainsOriginalInXlsxCarrier() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] sourceBytes = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string sourcePath = WriteTempWorkbook(sourceBytes, ".xls");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                ExcelDocumentConversionResult converted = ExcelDocument.Convert(
                    sourcePath,
                    destinationPath,
                    new ExcelDocumentConversionOptions {
                        CompatibilityMode = OfficeCompatibilityMode.PreservationOnly
                    });

                Assert.Contains(converted.Report.Diagnostics,
                    finding => finding.Code == "Excel.SourceCarrier.Embedded"
                        && finding.CompatibilityState == OfficeCompatibilityState.EmbeddedSource);
                Assert.Contains(converted.Report.Diagnostics,
                    finding => finding.RepresentsDataLoss
                        && finding.CompatibilityState == OfficeCompatibilityState.EmbeddedSource);
                using ExcelDocument modern = ExcelDocument.Load(destinationPath);
                Assert.True(modern.TryGetCompatibilitySourcePayload(out OfficeCompatibilitySourcePayload? payload, out string? error), error);
                Assert.NotNull(payload);
                Assert.Equal("Excel.Xls", payload!.FormatId);
                Assert.Equal(sourceBytes, payload.ToArray());

                using var packageStream = new MemoryStream(File.ReadAllBytes(destinationPath), writable: false);
                using ExcelDocument streamLoaded = ExcelDocument.Load(packageStream);
                Assert.True(streamLoaded.TryGetCompatibilitySourcePayload(
                    out OfficeCompatibilitySourcePayload? streamPayload,
                    out string? streamError), streamError);
                Assert.NotNull(streamPayload);
                Assert.Equal(sourceBytes, streamPayload!.ToArray());
            } finally {
                TryDelete(sourcePath);
                TryDelete(destinationPath);
            }
        }

        [Fact]
        public void LegacyXls_Convert_EditableAndVisualModesBlockUnmappedActiveContent() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] sourceBytes = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithVbaProjectStorage(workbookStream);
            string sourcePath = WriteTempWorkbook(sourceBytes, ".xls");

            try {
                foreach (OfficeCompatibilityMode mode in new[] {
                             OfficeCompatibilityMode.PreferEditable,
                             OfficeCompatibilityMode.PreferVisual
                         }) {
                    string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
                    var options = new ExcelDocumentConversionOptions { CompatibilityMode = mode };

                    ExcelDocumentConversionReport analysis = ExcelDocument.AnalyzeConversion(
                        sourcePath,
                        destinationPath,
                        options);

                    Assert.True(analysis.Compatibility.HasBlockedFeatures);
                    Assert.True(analysis.Compatibility.HasSecurityImpact);
                    Assert.Contains(analysis.Compatibility.Findings,
                        finding => finding.State == OfficeCompatibilityState.Blocked
                            && finding.RepresentsLoss
                            && (finding.Impact & OfficeCompatibilityImpact.Security) != 0);
                    ExcelDocumentConversionException blocked = Assert.Throws<ExcelDocumentConversionException>(() =>
                        ExcelDocument.Convert(sourcePath, destinationPath, options));
                    Assert.Equal(ExcelDocumentConversionFailureReason.DataLossBlocked, blocked.Reason);
                    Assert.False(File.Exists(destinationPath));
                }
            } finally {
                TryDelete(sourcePath);
            }
        }

        [Fact]
        public void LegacyXls_Convert_OleObjectStorageCarriesTypedSecurityImpact() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] sourceBytes = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithOleObjectStorage(workbookStream);
            string sourcePath = WriteTempWorkbook(sourceBytes, ".xls");
            string destinationPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                ExcelDocumentConversionReport analysis = ExcelDocument.AnalyzeConversion(
                    sourcePath,
                    destinationPath);

                Assert.True(analysis.Compatibility.HasSecurityImpact);
                Assert.Contains(analysis.Compatibility.Findings,
                    finding => finding.State == OfficeCompatibilityState.Blocked
                        && finding.RepresentsLoss
                        && (finding.Impact & OfficeCompatibilityImpact.Behavioral) != 0
                        && (finding.Impact & OfficeCompatibilityImpact.Security) != 0);
                ExcelDocumentConversionException blocked = Assert.Throws<ExcelDocumentConversionException>(() =>
                    ExcelDocument.Convert(sourcePath, destinationPath));
                Assert.Equal(ExcelDocumentConversionFailureReason.DataLossBlocked, blocked.Reason);
                Assert.False(File.Exists(destinationPath));
            } finally {
                TryDelete(sourcePath);
                TryDelete(destinationPath);
            }
        }

        [Fact]
        public void LegacyXls_Convert_DecryptedSourceReportsPasswordProtectionLoss() {
            const string password = "openpass";
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateRc4EncryptedWorkbookStream(password);
            byte[] sourceBytes = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string sourcePath = WriteTempWorkbook(sourceBytes, ".xls");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            var importOptions = new LegacyXlsImportOptions { Password = password };

            try {
                var blockedOptions = new ExcelDocumentConversionOptions {
                    LegacyXlsImportOptions = importOptions
                };
                ExcelDocumentConversionReport analysis = ExcelDocument.AnalyzeConversion(
                    sourcePath,
                    blockedPath,
                    blockedOptions);

                Assert.True(analysis.Compatibility.HasSecurityImpact);
                Assert.Contains(analysis.Compatibility.Findings,
                    finding => finding.Code == "Excel.PasswordEncryption.Removed"
                        && finding.State == OfficeCompatibilityState.Blocked
                        && finding.RepresentsLoss);
                ExcelDocumentConversionException blocked = Assert.Throws<ExcelDocumentConversionException>(() =>
                    ExcelDocument.Convert(sourcePath, blockedPath, blockedOptions));
                Assert.Equal(ExcelDocumentConversionFailureReason.DataLossBlocked, blocked.Reason);
                Assert.False(File.Exists(blockedPath));

                ExcelDocumentConversionResult allowed = ExcelDocument.Convert(
                    sourcePath,
                    allowedPath,
                    new ExcelDocumentConversionOptions {
                        CompatibilityMode = OfficeCompatibilityMode.BestEffort,
                        LegacyXlsImportOptions = new LegacyXlsImportOptions { Password = password }
                    });
                Assert.True(allowed.Report.Compatibility.HasSecurityImpact);
                Assert.Contains(allowed.Report.Compatibility.Findings,
                    finding => finding.Code == "Excel.PasswordEncryption.Removed"
                        && finding.State == OfficeCompatibilityState.Dropped
                        && finding.RepresentsLoss);
                Assert.True(File.Exists(allowedPath));
            } finally {
                TryDelete(sourcePath);
                TryDelete(blockedPath);
                TryDelete(allowedPath);
            }
        }
    }
}
