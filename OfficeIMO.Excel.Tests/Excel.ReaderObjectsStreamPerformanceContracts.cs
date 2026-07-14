using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private sealed class LargeStreamedRow {
            public int Id { get; set; }
            public string Name { get; set; } = string.Empty;
            public DateTime CreatedOn { get; set; }
            public double Amount { get; set; }
            public bool Active { get; set; }
        }

        [Fact]
        public void Reader_TypedObjectsStreamAndAutomaticMaterialization_MapLargeSortedRange() {
            const int dataRowCount = 4_096;
            var start = new DateTime(2025, 1, 1);
            using var package = new MemoryStream();

            using (var document = ExcelDocument.Create(package, new ExcelCreateOptions {
                PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose
            })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(1, 3, "CreatedOn");
                sheet.CellValue(1, 4, "Amount");
                sheet.CellValue(1, 5, "Active");

                for (int index = 1; index <= dataRowCount; index++) {
                    int row = index + 1;
                    sheet.CellValue(row, 1, index);
                    sheet.CellValue(row, 2, "Row " + index);
                    sheet.CellValue(row, 3, start.AddDays(index));
                    sheet.CellValue(row, 4, index + 0.25d);
                    sheet.CellValue(row, 5, (index & 1) == 0);
                }
            }

            using var reader = ExcelDocumentReader.Open(package.ToArray());
            var rows = reader.GetSheet("Data")
                .ReadObjectsStream<LargeStreamedRow>($"A1:E{dataRowCount + 1}")
                .ToList();

            Assert.Equal(dataRowCount, rows.Count);
            Assert.Equal(1, rows[0].Id);
            Assert.Equal("Row 2048", rows[2047].Name);
            var last = rows[dataRowCount - 1];
            Assert.Equal(start.AddDays(dataRowCount), last.CreatedOn);
            Assert.Equal(dataRowCount + 0.25d, last.Amount);
            Assert.True(last.Active);

            using var materializedReader = ExcelDocumentReader.Open(package.ToArray());
            var materializedRows = materializedReader.GetSheet("Data")
                .ReadObjects<LargeStreamedRow>($"A1:E{dataRowCount + 1}")
                .ToList();
            Assert.Equal(dataRowCount, materializedRows.Count);
            Assert.Equal("Row 2048", materializedRows[2047].Name);
            Assert.Equal(dataRowCount + 0.25d, materializedRows[dataRowCount - 1].Amount);
        }
    }
}
