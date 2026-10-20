using System.Data;
using OfficeIMO.Core.Internal.IO;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Tests;

public sealed class ExcelDataReaderOutputTests
{
    [Fact]
    public void HeldCanonicalClaimPreventsNoOverwritePublication()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-reader-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            var destination = Path.Combine(directory, "report.xlsx");
            using (var rows = CreateRows())
            using (var reader = rows.CreateDataReader())
            {
                using (OfficeFileCommitClaim.Acquire(destination))
                {
                    Assert.Throws<IOException>(() => ExcelDocumentService.WriteDataReaderPackage(
                        destination, reader, new ExcelTabularWriteOptions(), overwrite: false));
                    Assert.False(reader.IsClosed);
                    Assert.False(File.Exists(destination));
                }
            }

            Assert.Empty(Directory.GetFileSystemEntries(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void ExistingOutputIsRejectedBeforeConsumingTheCallerOwnedReader()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-reader-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            var destination = Path.Combine(directory, "report.xlsx");
            var original = new byte[] { 1, 2, 3, 4 };
            File.WriteAllBytes(destination, original);
            using (var rows = CreateRows())
            using (var reader = rows.CreateDataReader())
            {
                Assert.Throws<IOException>(() => ExcelDocumentService.WriteDataReaderPackage(
                    destination, reader, new ExcelTabularWriteOptions(), overwrite: false));
                Assert.False(reader.IsClosed);
                Assert.True(reader.Read());
                Assert.Equal("first", reader.GetString(0));
                Assert.Equal(7, reader.GetInt32(1));
            }

            Assert.Equal(original, File.ReadAllBytes(destination));
            Assert.Equal(new[] { destination }, Directory.GetFileSystemEntries(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void OverwritePublishesTheCompletedWorkbookAndLeavesTheReaderOpen()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-reader-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            var destination = Path.Combine(directory, "report.xlsx");
            File.WriteAllText(destination, "existing output");
            using (var rows = CreateRows())
            using (var reader = rows.CreateDataReader())
            {
                var result = ExcelDocumentService.WriteDataReaderPackage(
                    destination, reader, new ExcelTabularWriteOptions(), overwrite: true);
                Assert.Equal(1, result.RowCount);
                Assert.Equal(2, result.ColumnCount);
                Assert.False(reader.IsClosed);
            }

            using (var workbook = ExcelDocument.Load(destination))
            {
                var sheet = Assert.Single(workbook.Sheets);
                Assert.Equal("Name", sheet.CellAt(1, 1).GetValue<string>());
                Assert.Equal("first", sheet.CellAt(2, 1).GetValue<string>());
                Assert.Equal(7, sheet.CellAt(2, 2).GetValue<int>());
            }

            Assert.Equal(new[] { destination }, Directory.GetFileSystemEntries(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static DataTable CreateRows()
    {
        var rows = new DataTable();
        rows.Columns.Add("Name", typeof(string));
        rows.Columns.Add("Count", typeof(int));
        rows.Rows.Add("first", 7);
        return rows;
    }
}
