using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesContinuedSharedStringsAndLabelSstCells() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            string longAsciiText = new string('T', 32767);
            string longUnicodeText = new string('\u6f22', 5000);

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Long Strings");
                    sheet.CellValue(1, 1, longAsciiText);
                    sheet.CellValue(1, 2, longAsciiText);
                    sheet.CellValue(2, 1, longUnicodeText);
                    document.Save(xlsOutputPath);
                }

                IReadOnlyList<byte[]> sharedStringPayloads = GetBiffRecordPayloads(xlsOutputPath, 0x00fc);
                byte[] sharedStringPayload = Assert.Single(sharedStringPayloads);
                Assert.Equal(2U, ReadUInt32(sharedStringPayload, 4));
                Assert.Equal(3U, ReadUInt32(sharedStringPayload, 0));
                Assert.NotEmpty(GetBiffRecordPayloads(xlsOutputPath, 0x003c));
                Assert.Equal(3, GetBiffRecordPayloads(xlsOutputPath, 0x00fd).Count);
                Assert.Empty(GetBiffRecordPayloads(xlsOutputPath, 0x0204));
                byte[] extendedSharedStringPayload = Assert.Single(GetBiffRecordPayloads(xlsOutputPath, 0x00ff));
                Assert.Equal((ushort)8, ReadUInt16(extendedSharedStringPayload, 0));
                Assert.Equal((ushort)12, ReadUInt16(extendedSharedStringPayload, 6));
                Assert.True(ReadUInt32(extendedSharedStringPayload, 2) > ReadUInt16(extendedSharedStringPayload, 6));

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                Assert.Equal(longAsciiText, Assert.IsType<string>(Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1).Value));
                Assert.Equal(longAsciiText, Assert.IsType<string>(Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 2).Value));
                Assert.Equal(longUnicodeText, Assert.IsType<string>(Assert.Single(worksheet.Cells, cell => cell.Row == 2 && cell.Column == 1).Value));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesExtSstBucketsThatResolveToStringHeaders() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ExtSST");
                    for (int index = 0; index < 9; index++) {
                        int characterCount = index < 7 ? 1024 : index == 7 ? 1023 : 1008;
                        sheet.CellValue(index + 1, 1, new string((char)('A' + index), characterCount));
                    }

                    document.Save(xlsOutputPath);
                }

                byte[] workbookStream = ReadCompoundStream(File.ReadAllBytes(xlsOutputPath), "Workbook");
                byte[] extendedSharedStringPayload = Assert.Single(GetBiffRecordPayloads(xlsOutputPath, 0x00ff));
                Assert.Equal((ushort)8, ReadUInt16(extendedSharedStringPayload, 0));
                Assert.Equal(18, extendedSharedStringPayload.Length);

                AssertExtSstBucket(workbookStream, extendedSharedStringPayload, bucketIndex: 0, expectedRecordType: 0x00fc, expectedCharacterCount: 1024);
                AssertExtSstBucket(workbookStream, extendedSharedStringPayload, bucketIndex: 1, expectedRecordType: 0x003c, expectedCharacterCount: 1008);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesContinuedCachedFormulaStrings() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            string cachedText = new string('F', 32767);

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Formula String");
                    sheet.CellValue(1, 1, "Formula text");
                    Cell cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                        .Single(item => string.Equals(item.CellReference?.Value, "A1", StringComparison.OrdinalIgnoreCase));
                    cell.DataType = CellValues.String;
                    cell.CellFormula = new CellFormula("1");
                    cell.RemoveAllChildren<CellValue>();
                    cell.RemoveAllChildren<InlineString>();
                    cell.Append(new CellValue(cachedText));
                    sheet.WorksheetPart.Worksheet.Save();
                    document.Save(xlsOutputPath);
                }

                Assert.Single(GetBiffRecordPayloads(xlsOutputPath, 0x0207));
                Assert.NotEmpty(GetBiffRecordPayloads(xlsOutputPath, 0x003c));

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCell formulaCell = Assert.Single(worksheet.Cells);
                Assert.True(formulaCell.IsFormula);
                Assert.Equal(LegacyXlsCellValueKind.Text, formulaCell.Kind);
                Assert.Equal(cachedText, Assert.IsType<string>(formulaCell.Value));
                Assert.Equal("1", formulaCell.FormulaText);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        private static void AssertExtSstBucket(
            byte[] workbookStream,
            byte[] extendedSharedStringPayload,
            int bucketIndex,
            ushort expectedRecordType,
            ushort expectedCharacterCount) {
            int bucketOffset = checked(2 + (bucketIndex * 8));
            uint containingRecordOffset = ReadUInt32(extendedSharedStringPayload, bucketOffset);
            ushort stringRelativeOffset = ReadUInt16(extendedSharedStringPayload, bucketOffset + 4);
            int recordOffset = checked((int)containingRecordOffset);
            ushort recordLength = ReadUInt16(workbookStream, recordOffset + 2);

            Assert.Equal(expectedRecordType, ReadUInt16(workbookStream, recordOffset));
            Assert.InRange(stringRelativeOffset, (ushort)4, checked((ushort)(recordLength + 2)));
            Assert.Equal(expectedCharacterCount, ReadUInt16(workbookStream, checked(recordOffset + stringRelativeOffset)));
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesContinuedRichSharedStrings() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            string firstRunText = new string('A', 5000);
            string secondRunText = new string('B', 4000);

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Long Rich Text");
                    sheet.CellAt(1, 1).SetRichText(
                        new ExcelRichTextRun(firstRunText) { Bold = true },
                        new ExcelRichTextRun(secondRunText) { Italic = true });
                    document.Save(xlsOutputPath);
                }

                Assert.NotEmpty(GetBiffRecordPayloads(xlsOutputPath, 0x003c));
                Assert.Single(GetBiffRecordPayloads(xlsOutputPath, 0x00fd));

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsCell cell = Assert.Single(Assert.Single(result.Workbook.Worksheets).Cells);
                Assert.Equal(firstRunText + secondRunText, Assert.IsType<string>(cell.Value));
                Assert.Equal(2, cell.TextFormattingRuns.Count);
                Assert.Equal((ushort)0, cell.TextFormattingRuns[0].StartCharacter);
                Assert.Equal((ushort)5000, cell.TextFormattingRuns[1].StartCharacter);

                IReadOnlyList<ExcelRichTextRun> projectedRuns = result.Document.Sheets[0].CellAt(1, 1).GetRichText();
                Assert.Equal(2, projectedRuns.Count);
                Assert.Equal(firstRunText, projectedRuns[0].Text);
                Assert.True(projectedRuns[0].Bold);
                Assert.Equal(secondRunText, projectedRuns[1].Text);
                Assert.True(projectedRuns[1].Italic);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
