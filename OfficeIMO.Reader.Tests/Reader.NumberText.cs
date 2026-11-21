using System.Globalization;
using OfficeIMO.Excel;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderNumberTextTests {
    [Fact]
    public void ExcelStoredCellTextUsesShortestRoundTripNumbers() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Value");
                sheet.Cell(2, 1).SetValue(1D / 3D);
                document.Save();
            }

            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddExcelHandler(new ReaderExcelOptions { SheetName = "Data", A1Range = "A1:A2", CellText = ReaderExcelCellText.Raw })
                .Build();
            ReaderChunk chunk = Assert.Single(reader.Enumerate(path));

            Assert.Equal("0.3333333333333333", Assert.Single(Assert.Single(chunk.Tables!).Rows)[0]);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ChunkHashDistinguishesMetadataNumbersThatDifferAfterFifteenDigits() {
        double shortened = double.Parse("0.333333333333333", CultureInfo.InvariantCulture);
        Assert.NotEqual(1D / 3D, shortened);

        Assert.NotEqual(HashWithColumnConfidence(1D / 3D), HashWithColumnConfidence(shortened));
    }

    private static string HashWithColumnConfidence(double confidence) => DocumentReaderEngine.ComputeChunkHash(new ReaderChunk {
        Text = "Value",
        Tables = new[] {
            new ReaderTable {
                Columns = new[] { "Value" },
                Rows = new[] { new[] { "1" } },
                ColumnProfiles = new[] {
                    new ReaderTableColumnProfile { Index = 0, Name = "Value", Kind = ReaderTableColumnKind.Numeric, Confidence = confidence }
                }
            }
        }
    });
}
