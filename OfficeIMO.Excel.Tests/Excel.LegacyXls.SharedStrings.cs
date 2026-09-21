using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false, "XLS-BIFF-SST-SHORT")]
        [InlineData(true, "XLS-BIFF-SST-STRING-INVALID")]
        public void LegacyXls_Load_ClassifiesUnreadSharedStringsAsOmissions(bool invalidString, string code) {
            byte[] sst = invalidString
                ? new byte[] { 1, 0, 0, 0, 1, 0, 0, 0, 5, 0, 0, 0x41 }
                : new byte[] { 1, 0, 0, 0 };
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorkbookWithBrokenSst(sst);
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(
                new MemoryStream(compound), new LegacyXlsImportOptions { ReportUnsupportedContent = false });
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == code);
            Assert.Contains(result.FidelityDiagnostics, diagnostic =>
                diagnostic.Code == code && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.True(result.HasLoss);
            Assert.True(result.HasDocument);
            Assert.Throws<InvalidDataException>(() => result.RequireNoLoss());
            Assert.Contains(result.Summary.FidelityDiagnostics, diagnostic =>
                diagnostic.Code == code && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.Throws<InvalidDataException>(() => result.Summary.RequireNoLoss());
            LegacyXlsCell cell = Assert.Single(Assert.Single(result.AdvancedWorkbook.Worksheets).Cells);
            Assert.Equal(string.Empty, cell.Value);
        }

        [Theory]
        [InlineData(false, "XLS-BIFF-BOUNDSHEET-SHORT")]
        [InlineData(true, "XLS-BIFF-BOUNDSHEET-INVALID")]
        public void LegacyXls_Load_ClassifiesUnreadBoundSheetsAsOmissions(bool invalidName, string code) {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorkbookWithBrokenBoundSheet(invalidName);
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(
                new MemoryStream(compound), new LegacyXlsImportOptions { ReportUnsupportedContent = false });
            Assert.True(result.HasDocument);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == code);
            Assert.Contains(result.FidelityDiagnostics, diagnostic =>
                diagnostic.Code == code && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(result.Summary.FidelityDiagnostics, diagnostic =>
                diagnostic.Code == code && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.Single(result.AdvancedWorkbook.Worksheets);
            Assert.Throws<InvalidDataException>(() => result.RequireNoLoss());
            Assert.Throws<InvalidDataException>(() => result.Summary.RequireNoLoss());
        }

        [Fact]
        public void LegacyXls_Load_ClassifiesUnreadSheetStreamAsOmission() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorkbookWithInvalidSheetOffset();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(
                new MemoryStream(compound), new LegacyXlsImportOptions { ReportUnsupportedContent = false });
            Assert.True(result.HasDocument);
            Assert.Contains(result.FidelityDiagnostics, diagnostic =>
                diagnostic.Code == "XLS-BIFF-SHEET-OFFSET-INVALID"
                && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.Throws<InvalidDataException>(() => result.RequireNoLoss());
        }

        [Fact]
        public void LegacyXls_Load_ImportsSharedStringSplitInsideContinueRecord() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateSegmentedSharedStringWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-SST-STRING-INVALID");
            Assert.DoesNotContain(legacy.Diagnostics, d => d.RecordType == (ushort)BiffRecordType.Continue);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "AlphaBeta"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2 && Equals(cell.Value, "Second"));

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? first));
            Assert.Equal("AlphaBeta", first);
            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? second));
            Assert.Equal("Second", second);
        }

        [Fact]
        public void LegacyXls_Load_ImportsSharedStringWithRichAndExtendedPayloadsAsPlainText() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateRichExtendedSharedStringWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-SST-STRING-INVALID");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "RichText"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2 && Equals(cell.Value, "Plain"));

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? richText));
            Assert.Equal("RichText", richText);
            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? plainText));
            Assert.Equal("Plain", plainText);
        }

        [Fact]
        public void LegacyXls_Load_ImportsSharedStringWithRichRunsContinuedAfterCharacters() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateRichTextRunContinuedSharedStringWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-SST-STRING-INVALID");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "RichText"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2 && Equals(cell.Value, "Plain"));
        }

        private static partial class LegacyXlsTestWorkbookBuilder {
            internal static byte[] CreateWorkbookWithInvalidSheetOffset() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0, 6, 5, 0, 0xdb, 0x0b, 0xcc, 7 });
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(int.MaxValue, "Unread"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                return stream.ToArray();
            }

            internal static byte[] CreateWorkbookWithBrokenBoundSheet(bool invalidName) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0, 6, 5, 0, 0xdb, 0x0b, 0xcc, 7 });
                WriteRecord(stream, 0x0085, invalidName
                    ? new byte[] { 0, 0, 0, 0, 0, 0, 5 }
                    : new byte[] { 0 });
                long validBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Readable"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0, 6, 0x10, 0, 0xdb, 0x0b, 0xcc, 7 });
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)validBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateWorkbookWithBrokenSst(byte[] sst) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0, 6, 5, 0, 0xdb, 0x0b, 0xcc, 7 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Strings"));
                WriteRecord(stream, 0x00fc, sst);
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0, 6, 0x10, 0, 0xdb, 0x0b, 0xcc, 7 });
                WriteRecord(stream, 0x00fd, BuildLabelSstPayload(0, 0, 0));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateSegmentedSharedStringWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Strings"));
                WriteRecord(stream, 0x00fc, BuildSegmentedSstBasePayload("AlphaBeta", "Alpha", totalCount: 2, uniqueCount: 2));
                WriteRecord(stream, 0x003c, BuildSegmentedSstContinuePayload("Beta", "Second"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x00fd, BuildLabelSstPayload(0, 0, 0));
                WriteRecord(stream, 0x00fd, BuildLabelSstPayload(0, 1, 1));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateRichExtendedSharedStringWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "RichStrings"));
                WriteRecord(stream, 0x00fc, BuildRichExtendedSstPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x00fd, BuildLabelSstPayload(0, 0, 0));
                WriteRecord(stream, 0x00fd, BuildLabelSstPayload(0, 1, 1));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateRichTextRunContinuedSharedStringWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "RichRuns"));
                WriteRecord(stream, 0x00fc, BuildRichTextRunContinuedSstBasePayload("RichText"));
                WriteRecord(stream, 0x003c, BuildRichTextRunContinuedSstContinuePayload("Plain"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x00fd, BuildLabelSstPayload(0, 0, 0));
                WriteRecord(stream, 0x00fd, BuildLabelSstPayload(0, 1, 1));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            private static byte[] BuildSegmentedSstBasePayload(string text, string firstSegment, uint totalCount, uint uniqueCount) {
                using var stream = new MemoryStream();
                WriteUInt32(stream, totalCount);
                WriteUInt32(stream, uniqueCount);
                WriteUInt16(stream, checked((ushort)text.Length));
                stream.WriteByte(0);
                byte[] segmentBytes = Encoding.ASCII.GetBytes(firstSegment);
                stream.Write(segmentBytes, 0, segmentBytes.Length);
                return stream.ToArray();
            }

            private static byte[] BuildSegmentedSstContinuePayload(string continuedText, string nextString) {
                using var stream = new MemoryStream();
                stream.WriteByte(0);
                byte[] continuedBytes = Encoding.ASCII.GetBytes(continuedText);
                stream.Write(continuedBytes, 0, continuedBytes.Length);
                WriteSharedStringEntry(stream, nextString);
                return stream.ToArray();
            }

            private static byte[] BuildRichExtendedSstPayload() {
                using var stream = new MemoryStream();
                WriteUInt32(stream, 2);
                WriteUInt32(stream, 2);
                WriteRichExtendedSharedStringEntry(stream, "RichText");
                WriteSharedStringEntry(stream, "Plain");
                return stream.ToArray();
            }

            private static byte[] BuildRichTextRunContinuedSstBasePayload(string text) {
                byte[] textBytes = Encoding.ASCII.GetBytes(text);
                using var stream = new MemoryStream();
                WriteUInt32(stream, 2);
                WriteUInt32(stream, 2);
                WriteUInt16(stream, checked((ushort)textBytes.Length));
                stream.WriteByte(0x08);
                WriteUInt16(stream, 1);
                stream.Write(textBytes, 0, textBytes.Length);
                return stream.ToArray();
            }

            private static byte[] BuildRichTextRunContinuedSstContinuePayload(string nextString) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 1);
                WriteSharedStringEntry(stream, nextString);
                return stream.ToArray();
            }

            private static void WriteRichExtendedSharedStringEntry(Stream stream, string text) {
                byte[] textBytes = Encoding.ASCII.GetBytes(text);
                WriteUInt16(stream, checked((ushort)textBytes.Length));
                stream.WriteByte(0x0c);
                WriteUInt16(stream, 2);
                WriteUInt32(stream, 4);
                stream.Write(textBytes, 0, textBytes.Length);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 4);
                WriteUInt16(stream, 2);
                WriteUInt32(stream, 0);
            }
        }
    }
}
