using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelNamedRangesTests {
        [Fact]
        public void CanCreateAndReadNamedRanges() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                document.SetNamedRange("GlobalRange", "'Data'!A1:A2", save: false);
                sheet.SetNamedRange("LocalRange", "A1", save: false);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                Assert.Equal("'Data'!$A$1:$A$2", document.GetNamedRange("GlobalRange"));
                var sheet = document.Sheets.First(s => s.Name == "Data");
                Assert.Equal("$A$1", sheet.GetNamedRange("LocalRange"));
            }
            File.Delete(filePath);
        }

        [Fact]
        public void CanDeleteNamedRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.SetNamedRange("TempRange", "A1:B2", save: false);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets.First(s => s.Name == "Data");
                Assert.True(sheet.RemoveNamedRange("TempRange", save: false));
                Assert.Null(sheet.GetNamedRange("TempRange"));
                document.Save();
            }
            File.Delete(filePath);
        }

        [Fact]
        public void InvalidA1RangeThrows() {
            using var document = ExcelDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx"));
            document.AddWorkSheet("Data");
            Assert.Throws<ArgumentException>(() => document.SetNamedRange("Bad", "'Data'!A1:A"));
        }
    }
}

