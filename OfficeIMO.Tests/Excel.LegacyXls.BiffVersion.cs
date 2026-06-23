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
        }

        private static partial class LegacyXlsTestWorkbookBuilder {
            internal static byte[] CreateUnsupportedBiff5WorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x05, 0x05, 0x00 });
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
