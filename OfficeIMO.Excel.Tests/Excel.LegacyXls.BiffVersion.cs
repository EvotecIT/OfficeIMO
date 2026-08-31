using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_Load_AcceptsBiff5WorkbookGlobals() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateBiff5WorkbookGlobalsStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Empty(workbook.Worksheets);
            Assert.Empty(workbook.UnsupportedFeatures);
            Assert.DoesNotContain(workbook.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-VERSION-UNSUPPORTED");
            Assert.False(report.HasImportErrors);
            Assert.False(report.HasUnsupportedFeatures);
            Assert.False(report.UnsupportedFeaturesByKind.ContainsKey(LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion));
            Assert.Empty(report.UnsupportedFeaturesByDetail);
            Assert.Equal(0, report.PreservedFeatureRecordCount);
            Assert.False(report.PreservedFeatureRecordsByKind.ContainsKey(LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion));
            Assert.Equal(0, report.UnsupportedProjectionGapCount);
            Assert.Empty(report.UnsupportedProjectionGapsByKind);
            Assert.Equal(1, report.FileFormatStates["WorkbookFormat:SupportedBiff8"]);
            Assert.Equal(1, report.FileFormatStates["Encryption:Missing"]);
            Assert.Equal(1, report.FileFormatStates["UnsupportedBiffVersion:Missing"]);
            Assert.Equal(1, report.FileFormatStates["MalformedBof:Missing"]);
            Assert.Empty(report.FileFormatBlockers);
            Assert.Empty(report.FileFormatBlockersByRecordType);
            Assert.Empty(report.FileFormatBlockersByRecordName);
            Assert.Empty(report.FileFormatBlockersByLocation);
            Assert.False(report.UnsupportedBiffVersionsByVersion.ContainsKey("BIFF5"));
            Assert.Empty(report.UnsupportedBiffVersionsBySubstream);
            Assert.Empty(report.UnsupportedBiffVersionsByVersionAndSubstream);
            string markdown = report.ToMarkdown();
            Assert.Contains("File Format States", markdown);
            Assert.DoesNotContain("Unsupported BIFF Versions By Version", markdown);
        }

        [Fact]
        public void LegacyXls_Load_AcceptsBiff5WorksheetSubstream() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateBiff5WorksheetWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            LegacyXlsWorksheet sheet = Assert.Single(workbook.Worksheets);
            Assert.Equal("OldSheet", sheet.Name);
            LegacyXlsCell cell = Assert.Single(sheet.Cells);
            Assert.Equal(1, cell.Row);
            Assert.Equal(1, cell.Column);
            Assert.Equal("ShouldNotImport", cell.Value);
            Assert.Empty(workbook.UnsupportedFeatures);
            Assert.DoesNotContain(workbook.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-VERSION-UNSUPPORTED");
            LegacyXlsImportReport report = workbook.CreateImportReport();
            Assert.Equal(0, report.PreservedFeatureRecordCount);
            Assert.False(report.PreservedFeatureRecordsByKind.ContainsKey(LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion));
            Assert.Equal(0, report.UnsupportedProjectionGapCount);
            Assert.Empty(report.UnsupportedProjectionGapsByKind);
            Assert.Empty(report.FileFormatBlockers);
            Assert.Empty(report.FileFormatBlockersByRecordType);
            Assert.Empty(report.FileFormatBlockersByRecordName);
            Assert.Empty(report.FileFormatBlockersByLocation);
            Assert.False(report.UnsupportedBiffVersionsByVersion.ContainsKey("BIFF5"));
            Assert.Empty(report.UnsupportedBiffVersionsBySubstream);
            Assert.Empty(report.UnsupportedBiffVersionsByVersionAndSubstream);
        }

        [Fact]
        public void LegacyXls_GetSheetNames_ReadsBiff5WorkbookGlobalsOnly() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateBiff5WorkbookWithMalformedWorksheetBody();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string path = Path.Combine(
                Path.GetTempPath(),
                $"OfficeIMO.Excel.Biff5SheetNames.{Guid.NewGuid():N}.xls");
            try {
                File.WriteAllBytes(path, compound);

                Assert.Equal(
                    new[] { "OldSheet" },
                    OfficeIMO.Excel.ExcelDocument.GetSheetNames(path));
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void LegacyXls_Biff5SheetNames_UseWorkbookCodePage() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateBiff5WorkbookWithEncodedSheetName(
                new byte[] { 0xcb, 0xe8, 0xf1, 0xf2 },
                new byte[] { 0xe3, 0x04 });
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string path = Path.Combine(
                Path.GetTempPath(),
                $"OfficeIMO.Excel.Biff5CodePage.{Guid.NewGuid():N}.xls");
            try {
                File.WriteAllBytes(path, compound);

                Assert.Equal(new[] { "Лист" }, OfficeIMO.Excel.ExcelDocument.GetSheetNames(path));
                LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions());
                Assert.Equal((ushort)1251, workbook.CodePage);
                Assert.Equal("Лист", Assert.Single(workbook.Worksheets).Name);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void LegacyXls_Biff5SheetNames_SupportDbcsWorkbookCodePages() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateBiff5WorkbookWithEncodedSheetName(
                new byte[] { 0x83, 0x56, 0x81, 0x5b, 0x83, 0x67 },
                new byte[] { 0xa4, 0x03 });
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string path = Path.Combine(
                Path.GetTempPath(),
                $"OfficeIMO.Excel.Biff5DbcsCodePage.{Guid.NewGuid():N}.xls");
            try {
                File.WriteAllBytes(path, compound);

                Assert.Equal(new[] { "シート" }, OfficeIMO.Excel.ExcelDocument.GetSheetNames(path));
                LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions());
                Assert.Equal((ushort)932, workbook.CodePage);
                Assert.Equal("シート", Assert.Single(workbook.Worksheets).Name);
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void LegacyXls_Biff5SheetNames_RejectInvalidCodePageDeclarations(bool conflicting) {
            byte[][] codePagePayloads = conflicting
                ? new[] { new byte[] { 0xe3, 0x04 }, new byte[] { 0xe4, 0x04 } }
                : new[] { Array.Empty<byte>() };
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateBiff5WorkbookWithEncodedSheetName(
                new byte[] { 0xcb, 0xe8, 0xf1, 0xf2 },
                codePagePayloads);
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string path = Path.Combine(
                Path.GetTempPath(),
                $"OfficeIMO.Excel.Biff5InvalidCodePage.{Guid.NewGuid():N}.xls");
            try {
                File.WriteAllBytes(path, compound);

                Assert.Throws<InvalidDataException>(() => OfficeIMO.Excel.ExcelDocument.GetSheetNames(path));
                LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions());
                Assert.Empty(workbook.Worksheets);
                Assert.Contains(workbook.Diagnostics, diagnostic =>
                    diagnostic.Code == "XLS-BIFF-CODEPAGE-INVALID"
                    && diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void LegacyXls_Biff5SheetNames_AcceptDuplicateIdenticalCodePageDeclarations() {
            byte[] codePage = new byte[] { 0xe3, 0x04 };
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateBiff5WorkbookWithEncodedSheetName(
                new byte[] { 0xcb, 0xe8, 0xf1, 0xf2 },
                codePage,
                codePage);
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string path = Path.Combine(
                Path.GetTempPath(),
                $"OfficeIMO.Excel.Biff5DuplicateCodePage.{Guid.NewGuid():N}.xls");
            try {
                File.WriteAllBytes(path, compound);

                Assert.Equal(new[] { "Лист" }, OfficeIMO.Excel.ExcelDocument.GetSheetNames(path));
                LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions());
                Assert.Equal("Лист", Assert.Single(workbook.Worksheets).Name);
                Assert.DoesNotContain(workbook.Diagnostics, diagnostic =>
                    diagnostic.Code == "XLS-BIFF-CODEPAGE-INVALID");
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void LegacyXls_GetSheetNames_CountsAllBiff5SheetDefinitionsBeforeBodies() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateBiff5WorkbookWithChartDefinition();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string path = Path.Combine(
                Path.GetTempPath(),
                $"OfficeIMO.Excel.Biff5SheetLimit.{Guid.NewGuid():N}.xls");
            try {
                File.WriteAllBytes(path, compound);

                InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                    OfficeIMO.Excel.ExcelDocument.GetSheetNames(
                        path,
                        new OfficeIMO.Excel.ExcelReadOptions { MaxWorksheets = 1 }));
                Assert.Contains("more than the configured 1 worksheet definitions", exception.Message);
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void LegacyXls_GetSheetNames_RejectsMalformedBiff5SheetDefinitions(bool emptyName) {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateBiff5WorkbookWithMalformedBoundSheet(emptyName);
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string path = Path.Combine(
                Path.GetTempPath(),
                $"OfficeIMO.Excel.Biff5MalformedSheet.{Guid.NewGuid():N}.xls");
            try {
                File.WriteAllBytes(path, compound);

                Assert.Throws<InvalidDataException>(() =>
                    OfficeIMO.Excel.ExcelDocument.GetSheetNames(path));
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(0x0200, "BIFF2")]
        [InlineData(0x0300, "BIFF3")]
        [InlineData(0x0400, "BIFF4")]
        [InlineData(0x0700, "BIFF version 0x0700")]
        public void LegacyXls_Load_ReportsSpecificUnsupportedWorkbookBiffVersion(ushort version, string expectedVersionName) {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedBiffWorkbookStream(version);
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            string expectedDetail = $"BiffVersion:{expectedVersionName}:WorkbookGlobals";
            Assert.Empty(workbook.Worksheets);
            LegacyXlsUnsupportedFeature feature = Assert.Single(workbook.UnsupportedFeatures);
            Assert.Equal(LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion, feature.Kind);
            Assert.Equal(expectedDetail, feature.DetailCode);
            Assert.Contains(workbook.Diagnostics, diagnostic =>
                diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error
                && diagnostic.Code == "XLS-BIFF-VERSION-UNSUPPORTED"
                && diagnostic.DetailCode == expectedDetail);
            Assert.Equal(1, report.PreservedFeatureRecordCount);
            Assert.Equal(1, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion]);
            Assert.Equal(0, report.UnsupportedProjectionGapCount);
            Assert.Equal(1, report.FileFormatBlockersByRecordType["UnsupportedBiffVersion|0x0809"]);
            Assert.Equal(1, report.FileFormatBlockersByRecordName["UnsupportedBiffVersion|Record0x0809"]);
            Assert.Equal(1, report.FileFormatBlockersByLocation["XLS-BIFF-VERSION-UNSUPPORTED|(workbook)"]);
            Assert.Equal(1, report.UnsupportedBiffVersionsByVersion[expectedVersionName]);
            Assert.Equal(1, report.UnsupportedBiffVersionsBySubstream["WorkbookGlobals"]);
            Assert.Equal(1, report.UnsupportedBiffVersionsByVersionAndSubstream[$"{expectedVersionName}|WorkbookGlobals"]);
        }

        [Fact]
        public void LegacyXls_Load_ReportsMalformedBofFileFormatState() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorkbookWithMalformedBofStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Empty(workbook.Worksheets);
            Assert.Contains(workbook.Diagnostics, diagnostic =>
                diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error
                && diagnostic.Code == "XLS-BIFF-BOF-MISSING");
            Assert.Equal(1, report.FileFormatStates["WorkbookFormat:MalformedBof"]);
            Assert.Equal(1, report.FileFormatStates["MalformedBof:Present"]);
            Assert.Equal(1, report.FileFormatStates["Encryption:Missing"]);
            Assert.Equal(1, report.FileFormatStates["UnsupportedBiffVersion:Missing"]);
        }

        private static partial class LegacyXlsTestWorkbookBuilder {
            internal static byte[] CreateWorkbookWithMalformedBofStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                return stream.ToArray();
            }

            internal static byte[] CreateBiff5WorkbookGlobalsStream() {
                return CreateUnsupportedBiffWorkbookStream(0x0500);
            }

            internal static byte[] CreateUnsupportedBiffWorkbookStream(ushort version) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new[] { (byte)(version & 0x00ff), (byte)(version >> 8), (byte)0x05, (byte)0x00 });
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                return stream.ToArray();
            }

            internal static byte[] CreateBiff5WorksheetWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "OldSheet"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x05, 0x10, 0x00 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "ShouldNotImport"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateBiff5WorkbookWithMalformedWorksheetBody() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x05, 0x05, 0x00 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBiff5BoundSheetPayload(0, "OldSheet"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x05, 0x10, 0x00 });
                stream.WriteByte(0x04);
                stream.WriteByte(0x02);

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(
                    BitConverter.GetBytes(sheetOffset),
                    0,
                    bytes,
                    checked((int)boundSheetPosition + 4),
                    4);
                return bytes;
            }

            internal static byte[] CreateBiff5WorkbookWithEncodedSheetName(
                byte[] sheetNameBytes,
                params byte[][] codePagePayloads) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x05, 0x05, 0x00 });
                long boundSheetPosition = stream.Position;
                WriteRecord(
                    stream,
                    0x0085,
                    BuildBiff5BoundSheetPayload(
                        0,
                        sheetNameBytes));
                // Keep CodePage after BoundSheet to prove discovery is independent of record order.
                foreach (byte[] payload in codePagePayloads) {
                    WriteRecord(stream, 0x0042, payload);
                }
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x05, 0x10, 0x00 });
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(
                    BitConverter.GetBytes(sheetOffset),
                    0,
                    bytes,
                    checked((int)boundSheetPosition + 4),
                    4);
                return bytes;
            }

            internal static byte[] CreateBiff5WorkbookWithChartDefinition() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x05, 0x05, 0x00 });
                WriteRecord(stream, 0x0085, BuildBiff5BoundSheetPayload(0, "OldSheet"));
                WriteRecord(stream, 0x0085, BuildBiff5BoundSheetPayload(0, "Chart", sheetType: 2));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                return stream.ToArray();
            }

            internal static byte[] CreateBiff5WorkbookWithMalformedBoundSheet(bool emptyName) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x05, 0x05, 0x00 });
                WriteRecord(
                    stream,
                    0x0085,
                    emptyName
                        ? BuildBiff5BoundSheetPayload(0, string.Empty)
                        : new byte[6]);
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                return stream.ToArray();
            }

            private static byte[] BuildBiff5BoundSheetPayload(
                int streamOffset,
                string name,
                byte sheetType = 0) {
                return BuildBiff5BoundSheetPayload(
                    streamOffset,
                    System.Text.Encoding.ASCII.GetBytes(name),
                    sheetType);
            }

            private static byte[] BuildBiff5BoundSheetPayload(
                int streamOffset,
                byte[] nameBytes,
                byte sheetType = 0) {
                byte[] payload = new byte[7 + nameBytes.Length];
                Buffer.BlockCopy(BitConverter.GetBytes(streamOffset), 0, payload, 0, 4);
                payload[5] = sheetType;
                payload[6] = (byte)nameBytes.Length;
                Buffer.BlockCopy(nameBytes, 0, payload, 7, nameBytes.Length);
                return payload;
            }
        }
    }
}
