using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_SetCellValues_AdditionalTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "SetCellValuesAdditional.xlsx");
            var dateOffset = new DateTimeOffset(2024, 1, 2, 3, 4, 5, TimeSpan.Zero);
            var time = new TimeSpan(1, 2, 3, 4);
            uint ui = 123u;
            ulong ul = 1234567890UL;
            ushort us = 32100;
            byte by = 200;
            int? nullableInt = 42;
            int? nullableNull = null;
            DateTimeOffset? nullableDto = dateOffset;
            TimeSpan? nullableTs = time;

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, dateOffset);
                sheet.SetCellFormat(1, 1, "yyyy-mm-dd hh:mm");
                sheet.SetCellValue(2, 1, time);
                sheet.SetCellFormat(2, 1, "hh:mm:ss");
                sheet.SetCellValue(3, 1, ui);
                sheet.SetCellFormat(3, 1, "000000");
                sheet.SetCellValue(4, 1, ul);
                sheet.SetCellValue(5, 1, us);
                sheet.SetCellValue(6, 1, by);
                sheet.SetCellValue(7, 1, nullableInt);
                sheet.SetCellFormat(7, 1, "0");
                sheet.SetCellValue(8, 1, nullableNull);
                sheet.SetCellValue(9, 1, nullableDto);
                sheet.SetCellFormat(9, 1, "yyyy-mm-dd hh:mm");
                sheet.SetCellValue(10, 1, nullableTs);
                sheet.SetCellFormat(10, 1, "hh:mm:ss");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<Cell>().ToList();
                SharedStringTablePart shared = spreadsheet.WorkbookPart.SharedStringTablePart;

                Cell cellDto = cells.First(c => c.CellReference == "A1");
                Assert.Equal(dateOffset.UtcDateTime.ToOADate().ToString(CultureInfo.InvariantCulture), cellDto.CellValue.Text);

                Cell cellTs = cells.First(c => c.CellReference == "A2");
                Assert.Equal(time.TotalDays.ToString(CultureInfo.InvariantCulture), cellTs.CellValue.Text);

                Cell cellUint = cells.First(c => c.CellReference == "A3");
                Assert.Equal(ui.ToString(CultureInfo.InvariantCulture), cellUint.CellValue.Text);

                Cell cellUlong = cells.First(c => c.CellReference == "A4");
                Assert.Equal(((double)ul).ToString(CultureInfo.InvariantCulture), cellUlong.CellValue.Text);

                Cell cellUshort = cells.First(c => c.CellReference == "A5");
                Assert.Equal(us.ToString(CultureInfo.InvariantCulture), cellUshort.CellValue.Text);

                Cell cellByte = cells.First(c => c.CellReference == "A6");
                Assert.Equal(by.ToString(CultureInfo.InvariantCulture), cellByte.CellValue.Text);

                Cell cellNullableInt = cells.First(c => c.CellReference == "A7");
                Assert.Equal(nullableInt.Value.ToString(CultureInfo.InvariantCulture), cellNullableInt.CellValue.Text);

                Cell cellNullableNull = cells.First(c => c.CellReference == "A8");
                Assert.Equal(CellValues.SharedString, cellNullableNull.DataType.Value);
                Assert.Equal("0", cellNullableNull.CellValue.Text);
                Assert.NotNull(shared);
                Assert.Equal(string.Empty, shared.SharedStringTable.ElementAt(0).InnerText);

                Cell cellNullableDto = cells.First(c => c.CellReference == "A9");
                Assert.Equal(nullableDto.Value.UtcDateTime.ToOADate().ToString(CultureInfo.InvariantCulture), cellNullableDto.CellValue.Text);

                Cell cellNullableTs = cells.First(c => c.CellReference == "A10");
                Assert.Equal(nullableTs.Value.TotalDays.ToString(CultureInfo.InvariantCulture), cellNullableTs.CellValue.Text);

                var styles = spreadsheet.WorkbookPart.WorkbookStylesPart.Stylesheet;
                var numberingFormats = styles.NumberingFormats.Elements<NumberingFormat>().ToList();
                Assert.Contains(numberingFormats, n => n.FormatCode.Value == "yyyy-mm-dd hh:mm");
                Assert.Contains(numberingFormats, n => n.FormatCode.Value == "hh:mm:ss");
                Assert.Contains(numberingFormats, n => n.FormatCode.Value == "000000");
            }
        }
    }
}

