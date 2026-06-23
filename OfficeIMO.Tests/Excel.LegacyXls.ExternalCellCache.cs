using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_Load_PreservesExternalCellCacheRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateExternalCellCacheWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(workbook.Diagnostics, diagnostic =>
                diagnostic.Code == "XLS-BIFF-FEATURE-EXTERNAL-REFERENCE-UNSUPPORTED"
                && (diagnostic.RecordType == 0x0059 || diagnostic.RecordType == 0x005a));

            LegacyXlsExternalReference externalReference = Assert.Single(workbook.ExternalReferences);
            Assert.Equal(LegacyXlsExternalReferenceKind.ExternalWorkbook, externalReference.Kind);
            Assert.Equal("C:\\Data\\Budget.xls", externalReference.Target);

            LegacyXlsExternalCellCache cache = Assert.Single(externalReference.CachedCellCaches);
            Assert.True(cache.LinkValid);
            Assert.Equal(1, cache.DeclaredCrnCount);
            Assert.Equal(1, cache.SheetIndex);
            Assert.Equal("Feb", cache.SheetName);

            Assert.Collection(cache.Cells,
                cell => {
                    Assert.Equal(4, cell.Row);
                    Assert.Equal(0, cell.Column);
                    Assert.Equal(LegacyXlsCellValueKind.Number, cell.Kind);
                    Assert.Equal(12.5d, cell.Value);
                },
                cell => {
                    Assert.Equal(4, cell.Row);
                    Assert.Equal(1, cell.Column);
                    Assert.Equal(LegacyXlsCellValueKind.Text, cell.Kind);
                    Assert.Equal("Cached", cell.Value);
                },
                cell => {
                    Assert.Equal(4, cell.Row);
                    Assert.Equal(2, cell.Column);
                    Assert.Equal(LegacyXlsCellValueKind.Boolean, cell.Kind);
                    Assert.Equal(true, cell.Value);
                },
                cell => {
                    Assert.Equal(4, cell.Row);
                    Assert.Equal(3, cell.Column);
                    Assert.Equal(LegacyXlsCellValueKind.Error, cell.Kind);
                    Assert.Equal("#DIV/0!", cell.Value);
                },
                cell => {
                    Assert.Equal(4, cell.Row);
                    Assert.Equal(4, cell.Column);
                    Assert.Equal(LegacyXlsCellValueKind.Blank, cell.Kind);
                    Assert.Null(cell.Value);
                });

            Assert.Equal(1, report.ExternalReferenceCount);
            Assert.Equal(2, report.ExternalSheetNameCount);
            Assert.Equal(0, report.ExternalNameCount);
            Assert.Equal(1, report.ExternalCellCacheCount);
            Assert.Equal(5, report.ExternalCachedCellCount);
            Assert.Equal(1, report.ExternalReferencesByKind[LegacyXlsExternalReferenceKind.ExternalWorkbook]);
            Assert.Equal(1, report.ExternalReferencesByTarget["C:\\Data\\Budget.xls"]);
            Assert.Equal(2, report.ExternalSheetNamesByReferenceKind[LegacyXlsExternalReferenceKind.ExternalWorkbook]);
            Assert.Equal(1, report.ExternalCellCachesBySheetName["Feb"]);
            Assert.Equal(1, report.ExternalCachedCellsByValueKind[LegacyXlsCellValueKind.Blank]);
            Assert.Equal(1, report.ExternalCachedCellsByValueKind[LegacyXlsCellValueKind.Boolean]);
            Assert.Equal(1, report.ExternalCachedCellsByValueKind[LegacyXlsCellValueKind.Error]);
            Assert.Equal(1, report.ExternalCachedCellsByValueKind[LegacyXlsCellValueKind.Number]);
            Assert.Equal(1, report.ExternalCachedCellsByValueKind[LegacyXlsCellValueKind.Text]);
            string markdown = report.ToMarkdown();
            Assert.Contains("External References By Kind", markdown);
            Assert.Contains("External Cached Cells By Value Kind", markdown);
            Assert.Contains(workbook.UnsupportedFeatures, feature =>
                feature.Kind == LegacyXlsUnsupportedFeatureKind.ExternalReference
                && feature.DetailCode == "ExternalReference:ExternalWorkbook"
                && feature.RecordType == 0x01ae);
            Assert.DoesNotContain(workbook.UnsupportedFeatures, feature => feature.RecordType == 0x0059 || feature.RecordType == 0x005a);
        }

        private static partial class LegacyXlsTestWorkbookBuilder {
            internal static byte[] CreateExternalCellCacheWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Data"));
                WriteRecord(stream, 0x01ae, BuildSupBookExternalWorkbookPayload("C:\\Data\\Budget.xls", "Jan", "Feb"));
                WriteRecord(stream, 0x0059, BuildXctPayload(1, 1));
                WriteRecord(stream, 0x005a, BuildCrnPayload(4));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Local"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            private static byte[] BuildXctPayload(short crnCount, ushort sheetIndex) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, unchecked((ushort)crnCount));
                WriteUInt16(stream, sheetIndex);
                return stream.ToArray();
            }

            private static byte[] BuildCrnPayload(ushort row) {
                using var stream = new MemoryStream();
                stream.WriteByte(4);
                stream.WriteByte(0);
                WriteUInt16(stream, row);
                WriteSerNum(stream, 12.5d);
                WriteSerStr(stream, "Cached");
                WriteSerBool(stream, true);
                WriteSerErr(stream, 0x07);
                WriteSerNil(stream);
                return stream.ToArray();
            }

            private static void WriteSerNum(Stream stream, double value) {
                stream.WriteByte(0x01);
                byte[] bytes = BitConverter.GetBytes(value);
                stream.Write(bytes, 0, bytes.Length);
            }

            private static void WriteSerStr(Stream stream, string value) {
                stream.WriteByte(0x02);
                byte[] bytes = System.Text.Encoding.ASCII.GetBytes(value);
                WriteUInt16(stream, checked((ushort)value.Length));
                stream.WriteByte(0);
                stream.Write(bytes, 0, bytes.Length);
            }

            private static void WriteSerBool(Stream stream, bool value) {
                stream.WriteByte(0x04);
                stream.WriteByte(value ? (byte)1 : (byte)0);
                WriteIgnoredSerBytes(stream, 7);
            }

            private static void WriteSerErr(Stream stream, byte errorCode) {
                stream.WriteByte(0x10);
                stream.WriteByte(errorCode);
                WriteIgnoredSerBytes(stream, 7);
            }

            private static void WriteSerNil(Stream stream) {
                stream.WriteByte(0x00);
                WriteIgnoredSerBytes(stream, 8);
            }

            private static void WriteIgnoredSerBytes(Stream stream, int count) {
                for (int i = 0; i < count; i++) {
                    stream.WriteByte(0);
                }
            }
        }
    }
}
