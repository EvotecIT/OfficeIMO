using System.Text;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Excel;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderExcelSparseProjectionSecurityTests {
    [Fact]
    public void LegacySylkSparseRangeIsRejectedBeforeDenseProjection() {
        byte[] source = Encoding.ASCII.GetBytes("ID;PTEST\nC;X1;Y1;K1\nC;X16384;Y1001;K2\nE\n");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddLegacySpreadsheetHandler().Build();

        ReaderResourceLimitException error = Assert.Throws<ReaderResourceLimitException>(
            () => reader.ReadDocument(source, "sparse.slk"));

        Assert.Equal("excel.MaxProjectedCells", error.LimitName);
        Assert.Equal(1_000_000, error.Maximum);
    }

    [Fact]
    public void LegacySylkSmallSparseRangeStillProjectsItsValues() {
        byte[] source = Encoding.ASCII.GetBytes("ID;PTEST\nC;X1;Y1;K1\nC;X3;Y3;K2\nE\n");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddLegacySpreadsheetHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(source, "small.slk");

        Assert.Contains(result.Tables, table => table.Rows.Any(row => row.Contains("2")));
    }
}
