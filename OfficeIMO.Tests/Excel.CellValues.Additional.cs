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
        public void Test_CellValues_AdditionalTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesAdditional.xlsx");
            var dateOffset = new DateTimeOffset(2024, 1, 2, 3, 4, 5, TimeSpan.Zero);
            var time = new TimeSpan(1, 2, 3, 4);
#if NET6_0_OR_GREATER
            var dateOnly = new DateOnly(2024, 1, 2);
            var timeOnly = new TimeOnly(3, 4, 5);
#endif
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
                sheet.CellValue(1, 1, dateOffset);
                sheet.FormatCell(1, 1, "yyyy-mm-dd hh:mm");
                sheet.CellValue(2, 1, time);
                sheet.FormatCell(2, 1, "hh:mm:ss");
                sheet.CellValue(3, 1, ui);
                sheet.FormatCell(3, 1, "000000");
                sheet.CellValue(4, 1, ul);
                sheet.CellValue(5, 1, us);
                sheet.CellValue(6, 1, by);
                sheet.CellValue(7, 1, nullableInt);
                sheet.FormatCell(7, 1, "0");
                sheet.CellValue(8, 1, nullableNull);
                sheet.CellValue(9, 1, nullableDto);
                sheet.FormatCell(9, 1, "yyyy-mm-dd hh:mm");
                sheet.CellValue(10, 1, nullableTs);
                sheet.FormatCell(10, 1, "hh:mm:ss");
#if NET6_0_OR_GREATER
                sheet.CellValue(11, 1, dateOnly);
                sheet.FormatCell(11, 1, "yyyy-mm-dd");
                sheet.CellValue(12, 1, timeOnly);
                sheet.FormatCell(12, 1, "hh:mm:ss");
#endif
                document.Save();
            }

            SpreadsheetDocument spreadsheet = null!;
            Exception? ex = Record.Exception(() => spreadsheet = SpreadsheetDocument.Open(filePath, false));
            Assert.Null(ex);
            using (spreadsheet) {
                ValidateSpreadsheetDocument(filePath, spreadsheet);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<Cell>().ToList();
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;

                Cell cellDto = cells.First(c => c.CellReference == "A1");
                Assert.Equal(dateOffset.LocalDateTime.ToOADate().ToString(CultureInfo.InvariantCulture), cellDto.CellValue!.Text);

                Cell cellTs = cells.First(c => c.CellReference == "A2");
                Assert.Equal(time.TotalDays.ToString(CultureInfo.InvariantCulture), cellTs.CellValue!.Text);

                Cell cellUint = cells.First(c => c.CellReference == "A3");
                Assert.Equal(ui.ToString(CultureInfo.InvariantCulture), cellUint.CellValue!.Text);

                Cell cellUlong = cells.First(c => c.CellReference == "A4");
                Assert.Equal(((double)ul).ToString(CultureInfo.InvariantCulture), cellUlong.CellValue!.Text);

                Cell cellUshort = cells.First(c => c.CellReference == "A5");
                Assert.Equal(us.ToString(CultureInfo.InvariantCulture), cellUshort.CellValue!.Text);

                Cell cellByte = cells.First(c => c.CellReference == "A6");
                Assert.Equal(by.ToString(CultureInfo.InvariantCulture), cellByte.CellValue!.Text);

                Cell cellNullableInt = cells.First(c => c.CellReference == "A7");
                Assert.Equal(nullableInt.Value.ToString(CultureInfo.InvariantCulture), cellNullableInt.CellValue!.Text);

                Cell cellNullableNull = cells.First(c => c.CellReference == "A8");
                Assert.Equal(CellValues.SharedString, cellNullableNull.DataType!.Value);
                Assert.Equal("0", cellNullableNull.CellValue!.Text);
                Assert.NotNull(shared);
                Assert.Equal(string.Empty, shared!.SharedStringTable!.ElementAt(0).InnerText);

                Cell cellNullableDto = cells.First(c => c.CellReference == "A9");
                Assert.Equal(nullableDto.Value.LocalDateTime.ToOADate().ToString(CultureInfo.InvariantCulture), cellNullableDto.CellValue!.Text);

                Cell cellNullableTs = cells.First(c => c.CellReference == "A10");
                Assert.Equal(nullableTs.Value.TotalDays.ToString(CultureInfo.InvariantCulture), cellNullableTs.CellValue!.Text);
#if NET6_0_OR_GREATER
                Cell cellDateOnly = cells.First(c => c.CellReference == "A11");
                Assert.Equal(dateOnly.ToDateTime(TimeOnly.MinValue).ToOADate().ToString(CultureInfo.InvariantCulture), cellDateOnly.CellValue!.Text);

                Cell cellTimeOnly = cells.First(c => c.CellReference == "A12");
                Assert.Equal(timeOnly.ToTimeSpan().TotalDays.ToString(CultureInfo.InvariantCulture), cellTimeOnly.CellValue!.Text);
#endif

                var styles = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                var numberingFormats = styles.NumberingFormats!.Elements<NumberingFormat>().ToList();
                Assert.Contains(numberingFormats, n => n.FormatCode != null && n.FormatCode.Value == "yyyy-mm-dd hh:mm");
                Assert.Contains(numberingFormats, n => n.FormatCode != null && n.FormatCode.Value == "hh:mm:ss");
                Assert.Contains(numberingFormats, n => n.FormatCode != null && n.FormatCode.Value == "000000");
#if NET6_0_OR_GREATER
                Assert.Contains(numberingFormats, n => n.FormatCode != null && n.FormatCode.Value == "yyyy-mm-dd");
#endif
            }
        }

        [Fact]
        public void Test_CellValues_LargeIntegers_AreSerializedExactly() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesLargeIntegers.xlsx");
            ulong largeUnsigned = ulong.MaxValue;
            long largeSigned = long.MinValue;

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, largeUnsigned);
                sheet.CellValue(2, 1, (object)largeSigned);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                ValidateSpreadsheetDocument(filePath, spreadsheet);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<Cell>().ToList();

                Cell cellLargeUnsigned = cells.First(c => c.CellReference == "A1");
                Assert.Equal(CellValues.Number, cellLargeUnsigned.DataType!.Value);
                Assert.Equal(largeUnsigned.ToString(CultureInfo.InvariantCulture), cellLargeUnsigned.CellValue!.Text);

                Cell cellLargeSigned = cells.First(c => c.CellReference == "A2");
                Assert.Equal(CellValues.Number, cellLargeSigned.DataType!.Value);
                Assert.Equal(largeSigned.ToString(CultureInfo.InvariantCulture), cellLargeSigned.CellValue!.Text);
            }
        }

        [Fact]
        public void Test_CellValues_DateTimeOffset_RoundTrip_LocalTimes()
        {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesDateTimeOffsetRoundTrip.xlsx");
            var positive = new DateTimeOffset(2024, 5, 1, 10, 30, 0, TimeSpan.FromHours(9));
            var negative = new DateTimeOffset(2024, 5, 1, 10, 30, 0, TimeSpan.FromHours(-4));

            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, positive);
                sheet.CellValue(2, 1, negative);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                ValidateSpreadsheetDocument(filePath, spreadsheet);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<Cell>().ToList();

                Cell cellPositive = cells.First(c => c.CellReference == "A1");
                Cell cellNegative = cells.First(c => c.CellReference == "A2");

                double positiveSerial = double.Parse(cellPositive.CellValue!.Text, CultureInfo.InvariantCulture);
                double negativeSerial = double.Parse(cellNegative.CellValue!.Text, CultureInfo.InvariantCulture);

                Assert.Equal(positive.LocalDateTime, DateTime.FromOADate(positiveSerial));
                Assert.Equal(negative.LocalDateTime, DateTime.FromOADate(negativeSerial));
            }
        }

        [Fact]
        public void Test_CellValues_DateTimeOffset_CustomStrategy_UsesProvidedDelegate()
        {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesDateTimeOffsetUtc.xlsx");
            var value = new DateTimeOffset(2024, 5, 1, 10, 30, 0, TimeSpan.FromHours(2));

            using (var document = ExcelDocument.Create(filePath))
            {
                document.DateTimeOffsetWriteStrategy = dto => dto.UtcDateTime;
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, value);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                ValidateSpreadsheetDocument(filePath, spreadsheet);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cell = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");

                Assert.Equal(value.UtcDateTime.ToOADate().ToString(CultureInfo.InvariantCulture), cell.CellValue!.Text);
            }
        }

        [Fact]
        public void Test_CellValues_DateTimeOffset_BeforeExcelEpoch_FallsBackToSharedString()
        {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesDateTimeOffsetBeforeEpoch.xlsx");
            var historical = new DateTimeOffset(1899, 12, 31, 23, 59, 0, TimeSpan.Zero);

            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, historical);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                ValidateSpreadsheetDocument(filePath, spreadsheet);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cell = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");

                Assert.Equal(CellValues.SharedString, cell.DataType!.Value);

                var sharedTable = spreadsheet.WorkbookPart!.SharedStringTablePart!;
                int index = int.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture);
                var text = sharedTable.SharedStringTable!.ElementAt(index).InnerText;

                Assert.Equal(historical.ToString("o", CultureInfo.InvariantCulture), text);
            }
        }

        [Fact]
        public void Test_CellValues_DateTimeOffset_StrategyExceptionWrapped()
        {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesDateTimeOffsetStrategyException.xlsx");
            var value = new DateTimeOffset(2024, 5, 1, 10, 30, 0, TimeSpan.FromHours(1));

            using (var document = ExcelDocument.Create(filePath))
            {
                document.DateTimeOffsetWriteStrategy = _ => throw new InvalidOperationException("boom");
                var sheet = document.AddWorkSheet("Data");

                var ex = Assert.Throws<InvalidOperationException>(() => sheet.CellValue(1, 1, value));
                Assert.Contains("DateTimeOffset write strategy", ex.Message);
                Assert.IsType<InvalidOperationException>(ex.InnerException);
            }
        }

        [Fact]
        public void Test_CellValues_TimeSpanFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesTimeSpanFormat.xlsx");
            var duration = new TimeSpan(1, 2, 3, 4);

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, duration);
                sheet.CellValue(2, 1, (TimeSpan?)duration);
                document.Save();
            }

            SpreadsheetDocument spreadsheet = null!;
            Exception? ex = Record.Exception(() => spreadsheet = SpreadsheetDocument.Open(filePath, false));
            Assert.Null(ex);

            using (spreadsheet) {
                ValidateSpreadsheetDocument(filePath, spreadsheet);

                var workbookPart = spreadsheet.WorkbookPart!;
                Assert.NotNull(workbookPart.WorkbookStylesPart);

                WorksheetPart wsPart = workbookPart.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<Cell>().ToList();

                Cell cellDuration = cells.First(c => c.CellReference == "A1");
                Assert.Equal(CellValues.Number, cellDuration.DataType!.Value);
                Assert.Equal(duration.TotalDays.ToString(CultureInfo.InvariantCulture), cellDuration.CellValue!.Text);
                Assert.NotNull(cellDuration.StyleIndex);

                var styles = workbookPart.WorkbookStylesPart!.Stylesheet!;
                Assert.NotNull(styles.CellFormats);
                var cellFormats = styles.CellFormats!.Elements<CellFormat>().ToList();
                var cellFormat = cellFormats[(int)cellDuration.StyleIndex!.Value];
                Assert.NotNull(cellFormat.NumberFormatId);
                Assert.Equal(46U, cellFormat.NumberFormatId!.Value);
                Assert.True(cellFormat.ApplyNumberFormat?.Value ?? false);

                Cell cellNullable = cells.First(c => c.CellReference == "A2");
                Assert.Equal(duration.TotalDays.ToString(CultureInfo.InvariantCulture), cellNullable.CellValue!.Text);
                Assert.NotNull(cellNullable.StyleIndex);
                Assert.Equal(cellDuration.StyleIndex!.Value, cellNullable.StyleIndex!.Value);
            }
        }
    }
}

