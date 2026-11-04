using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Reader_CustomNumericFormatWithQuotedText_RemainsNumeric() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderQuotedTextNumberFormat.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("Value");
                    sheet.Cell(2, 1).SetValue(3d);
                    sheet.ColumnStyleByHeader("Value").NumberFormat("0 \"days\"");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var value = reader.GetSheet("Data").EnumerateCells().Single(c => c.Row == 2 && c.Column == 1).Value;

                Assert.IsType<double>(value);
                Assert.Equal(3d, (double)value!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_DurationFormatWithBracketHours_StillReadsAsDateLike() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderBracketHoursNumberFormat.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("Duration");
                    sheet.Cell(2, 1).SetValue(1.5d);
                    sheet.ColumnStyleByHeader("Duration").NumberFormat("[h]:mm");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var value = reader.GetSheet("Data").EnumerateCells().Single(c => c.Row == 2 && c.Column == 1).Value;

                Assert.IsType<DateTime>(value);
                Assert.Equal(DateTime.FromOADate(1.5d), (DateTime)value!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_CustomNumericFormatWithEscapedHourLiteral_RemainsNumeric() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderEscapedHourLiteralNumberFormat.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("Hours");
                    sheet.Cell(2, 1).SetValue(7d);
                    sheet.ColumnStyleByHeader("Hours").NumberFormat("0\\h");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var value = reader.GetSheet("Data").EnumerateCells().Single(c => c.Row == 2 && c.Column == 1).Value;

                Assert.IsType<double>(value);
                Assert.Equal(7d, (double)value!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
