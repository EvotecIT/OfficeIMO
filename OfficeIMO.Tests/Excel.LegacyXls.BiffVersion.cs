using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_Load_ReportsUnsupportedWorkbookBiffVersion() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedBiff5WorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Empty(workbook.Worksheets);
            LegacyXlsUnsupportedFeature feature = Assert.Single(workbook.UnsupportedFeatures);
            Assert.Equal(LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion, feature.Kind);
            Assert.Equal("XLS-BIFF-VERSION-UNSUPPORTED", feature.Code);
            Assert.Equal("BiffVersion:BIFF5:WorkbookGlobals", feature.DetailCode);
            Assert.Contains(workbook.Diagnostics, diagnostic =>
                diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error
                && diagnostic.Code == "XLS-BIFF-VERSION-UNSUPPORTED"
                && diagnostic.DetailCode == "BiffVersion:BIFF5:WorkbookGlobals");
            Assert.True(report.HasImportErrors);
            Assert.True(report.HasUnsupportedFeatures);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["UnsupportedBiffVersion|XLS-BIFF-VERSION-UNSUPPORTED|BiffVersion:BIFF5:WorkbookGlobals"]);
            Assert.Equal(1, report.FileFormatBlockers["UnsupportedBiffVersion|BiffVersion:BIFF5:WorkbookGlobals"]);
            Assert.Equal(1, report.UnsupportedBiffVersionsByVersion["BIFF5"]);
            Assert.Equal(1, report.UnsupportedBiffVersionsBySubstream["WorkbookGlobals"]);
            Assert.Equal(1, report.UnsupportedBiffVersionsByVersionAndSubstream["BIFF5|WorkbookGlobals"]);
            string markdown = report.ToMarkdown();
            Assert.Contains("File Format Blockers", markdown);
            Assert.Contains("Unsupported BIFF Versions By Version", markdown);
            Assert.Contains("Unsupported BIFF Versions By Substream", markdown);
            Assert.Contains("Unsupported BIFF Versions By Version And Substream", markdown);
        }

        [Fact]
        public void LegacyXls_Load_ReportsUnsupportedWorksheetBiffVersion() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedWorksheetBiff5WorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            LegacyXlsWorksheet sheet = Assert.Single(workbook.Worksheets);
            Assert.Equal("OldSheet", sheet.Name);
            Assert.Empty(sheet.Cells);
            LegacyXlsUnsupportedFeature feature = Assert.Single(workbook.UnsupportedFeatures);
            Assert.Equal(LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion, feature.Kind);
            Assert.Equal("OldSheet", feature.SheetName);
            Assert.Equal("BiffVersion:BIFF5:Worksheet", feature.DetailCode);
            Assert.Contains(workbook.Diagnostics, diagnostic =>
                diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error
                && diagnostic.SheetName == "OldSheet"
                && diagnostic.Code == "XLS-BIFF-VERSION-UNSUPPORTED"
                && diagnostic.DetailCode == "BiffVersion:BIFF5:Worksheet");
            LegacyXlsImportReport report = workbook.CreateImportReport();
            Assert.Equal(1, report.FileFormatBlockers["UnsupportedBiffVersion|BiffVersion:BIFF5:Worksheet"]);
            Assert.Equal(1, report.UnsupportedBiffVersionsByVersion["BIFF5"]);
            Assert.Equal(1, report.UnsupportedBiffVersionsBySubstream["Worksheet"]);
            Assert.Equal(1, report.UnsupportedBiffVersionsByVersionAndSubstream["BIFF5|Worksheet"]);
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
                ReportUnsupportedRecords = true
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
            Assert.Equal(1, report.UnsupportedBiffVersionsByVersion[expectedVersionName]);
            Assert.Equal(1, report.UnsupportedBiffVersionsBySubstream["WorkbookGlobals"]);
            Assert.Equal(1, report.UnsupportedBiffVersionsByVersionAndSubstream[$"{expectedVersionName}|WorkbookGlobals"]);
        }

        private static partial class LegacyXlsTestWorkbookBuilder {
            internal static byte[] CreateUnsupportedBiff5WorkbookStream() {
                return CreateUnsupportedBiffWorkbookStream(0x0500);
            }

            internal static byte[] CreateUnsupportedBiffWorkbookStream(ushort version) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new[] { (byte)(version & 0x00ff), (byte)(version >> 8), (byte)0x05, (byte)0x00 });
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                return stream.ToArray();
            }

            internal static byte[] CreateUnsupportedWorksheetBiff5WorkbookStream() {
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
