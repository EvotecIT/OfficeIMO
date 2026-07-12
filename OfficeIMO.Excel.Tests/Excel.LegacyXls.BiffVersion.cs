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
        }
    }
}
