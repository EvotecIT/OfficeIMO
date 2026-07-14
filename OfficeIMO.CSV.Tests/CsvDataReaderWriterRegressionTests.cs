using System.Data;
using System.Globalization;
using System.IO;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvDataReaderWriterRegressionTests
{
    [Theory]
    [InlineData('T', "\"True\"TAlpha\nFalseTBeta\n")]
    [InlineData('F', "TrueFAlpha\n\"False\"FBeta\n")]
    public void WriteDataReader_QuotesBooleanWhenDelimiterAppearsInLiteral(char delimiter, string expected)
    {
        var table = new DataTable();
        table.Columns.Add("Enabled", typeof(bool));
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add(true, "Alpha");
        table.Rows.Add(false, "Beta");

        using var reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        CsvDocument.WriteDataReader(
            writer,
            reader,
            new CsvSaveOptions { Delimiter = delimiter, IncludeHeader = false, NewLine = "\n" });

        Assert.Equal(expected, writer.ToString());
    }
}
