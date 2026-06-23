using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_Load_ImportsPhase3StyleInheritance() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase3StyleInheritanceWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(2, legacy.CellFormats.Count);
            Assert.True(legacy.CellFormats[0].IsStyle);
            Assert.False(legacy.CellFormats[1].IsStyle);
            Assert.Equal(0, legacy.CellFormats[1].ParentStyleIndex);
            Assert.False(legacy.CellFormats[1].ApplyNumberFormat);
            Assert.False(legacy.CellFormats[1].ApplyFont);
            Assert.False(legacy.CellFormats[1].ApplyAlignment);
            Assert.Null(legacy.CellFormats[1].Border);
            Assert.Equal(0, legacy.CellFormats[1].FillPattern);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
            Dictionary<string, Cell> cells = worksheetPart.Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);

            Cell inheritedCell = cells["A1"];
            Assert.Equal(new DateTime(2024, 2, 3).ToOADate().ToString(CultureInfo.InvariantCulture), inheritedCell.CellValue!.Text);
            Assert.NotNull(inheritedCell.StyleIndex);
            CellFormat inheritedFormat = workbookPart.WorkbookStylesPart!.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)inheritedCell.StyleIndex!.Value);

            NumberingFormat inheritedNumberFormat = Assert.Single(workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats!.Elements<NumberingFormat>(),
                format => format.NumberFormatId!.Value == inheritedFormat.NumberFormatId!.Value);
            Assert.Equal("yyyy-mm-dd", inheritedNumberFormat.FormatCode!.Value);

            DocumentFormat.OpenXml.Spreadsheet.Font inheritedFont = workbookPart.WorkbookStylesPart.Stylesheet.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt((int)inheritedFormat.FontId!.Value);
            Assert.Equal("Consolas", inheritedFont.FontName!.Val!.Value);
            Assert.Equal(13d, inheritedFont.FontSize!.Val!.Value);
            Assert.Equal("FF123456", inheritedFont.Color!.Rgb!.Value);
            Assert.NotNull(inheritedFont.Bold);

            Fill inheritedFill = workbookPart.WorkbookStylesPart.Stylesheet.Fills!.Elements<Fill>().ElementAt((int)inheritedFormat.FillId!.Value);
            Assert.Equal(PatternValues.Solid, inheritedFill.PatternFill!.PatternType!.Value);
            Assert.Equal("FF123456", inheritedFill.PatternFill.ForegroundColor!.Rgb!.Value);

            Assert.True(inheritedFormat.ApplyAlignment!.Value);
            Assert.Equal(HorizontalAlignmentValues.Center, inheritedFormat.Alignment!.Horizontal!.Value);
            Assert.Equal(VerticalAlignmentValues.Center, inheritedFormat.Alignment.Vertical!.Value);
            Assert.True(inheritedFormat.Alignment.WrapText!.Value);

            Assert.True(inheritedFormat.ApplyBorder!.Value);
            Border inheritedBorder = workbookPart.WorkbookStylesPart.Stylesheet.Borders!.Elements<Border>().ElementAt((int)inheritedFormat.BorderId!.Value);
            Assert.Equal(BorderStyleValues.Thin, inheritedBorder.LeftBorder!.Style!.Value);
            Assert.Equal("FFABCDEF", inheritedBorder.LeftBorder.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>()!.Rgb!.Value);

            Assert.True(inheritedFormat.ApplyProtection!.Value);
            Assert.False(inheritedFormat.Protection!.Locked!.Value);
            Assert.True(inheritedFormat.Protection.Hidden!.Value);
        }

        private static partial class LegacyXlsTestWorkbookBuilder {
            internal static byte[] CreatePhase3StyleInheritanceWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Inherited"));
                WriteRecord(stream, 0x0031, BuildFontPayload("Arial", 11d, bold: false, italic: false, underline: false));
                WriteRecord(stream, 0x0031, BuildFontPayload("Arial", 11d, bold: true, italic: false, underline: false));
                WriteRecord(stream, 0x0031, BuildFontPayload("Arial", 11d, bold: false, italic: true, underline: false));
                WriteRecord(stream, 0x0031, BuildFontPayload("Arial", 11d, bold: true, italic: true, underline: false));
                WriteRecord(stream, 0x0031, BuildFontPayload("Consolas", 13d, bold: true, italic: false, underline: false, colorIndex: 0x0008));
                WriteRecord(stream, 0x0092, BuildPalettePayload("FF123456", "FFABCDEF"));
                WriteRecord(stream, 0x041e, BuildFormatPayload(164, "yyyy-mm-dd"));
                WriteRecord(stream, 0x00e0, BuildXfPayload(
                    164,
                    fontIndex: 5,
                    isStyle: true,
                    parentStyleIndex: 0x0fff,
                    fillPattern: 1,
                    fillForegroundColorIndex: 0x0008,
                    applyAlignment: true,
                    horizontalAlignment: 2,
                    verticalAlignment: 1,
                    wrapText: true,
                    locked: false,
                    formulaHidden: true,
                    applyProtection: true,
                    leftBorderStyle: 1,
                    leftBorderColorIndex: 0x0009));
                WriteRecord(stream, 0x00e0, BuildXfPayload(
                    0,
                    parentStyleIndex: 0,
                    applyNumberFormat: false,
                    applyFont: false));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, new DateTime(2024, 2, 3).ToOADate(), styleIndex: 1));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }
        }
    }
}
