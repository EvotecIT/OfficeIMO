using System.Globalization;
using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

/// <summary>
/// A schema read from a byte stream returns the same typed values, missing values and conversion
/// failures as the same schema over decoded text.
/// </summary>
public sealed class CsvSchemaStreamingParityTests {
    private const string Input =
        "Id,Name,Price,When,Active,Note\n" +
        "1,Ada,12.50,2026-01-02,true,plain\n" +
        "2,\"Bob, Jr.\",\"7\",2026-02-03,false,\"say \"\"hi\"\"\"\n" +
        "x,Cy,abc,not-a-date,maybe,\n" +
        ",,,,,\n" +
        "5,Zoë,1e3,2026-03-04T05:06:07,1\n" +
        "6,Eve,3.25,2026-12-31,0,\"two\nlines\"\n";

    [Theory]
    [InlineData(DataMappingErrorValuePolicy.Include, false, false)]
    [InlineData(DataMappingErrorValuePolicy.Include, true, false)]
    [InlineData(DataMappingErrorValuePolicy.Redact, false, false)]
    [InlineData(DataMappingErrorValuePolicy.Include, false, true)]
    [InlineData(DataMappingErrorValuePolicy.Redact, true, true)]
    public void SchemaValuesMatchDecodedText(DataMappingErrorValuePolicy policy, bool trim, bool generalColumn) {
        Assert.Equal(Observe(reader: null, policy, trim, generalColumn), Observe(reader: Input, policy, trim, generalColumn));
    }

    [Theory]
    [InlineData(1000, null)]
    [InlineData(2, null)]
    [InlineData(1000, "x")]
    public void InferredSchemaMatchesDecodedText(int sampleRows, string? nullValue) {
        Assert.Equal(Observe(reader: null, sampleRows, nullValue), Observe(reader: Input, sampleRows, nullValue));
    }

    private static string Observe(string? reader, int sampleRows, string? nullValue) =>
        Read(reader, new CsvReaderOptions { InferSchema = true, SchemaSampleRows = sampleRows, NullValue = nullValue });

    // A general column (a decimal with a default) converts through values; the others convert text directly.
    private static string Observe(string? reader, DataMappingErrorValuePolicy policy, bool trim, bool generalColumn) {
        var builder = new CsvSchemaBuilder();
        builder.Column("Id").AsInt32().Required();
        builder.Column("Name").AsString().Required();
        if (generalColumn) builder.Column("Price").AsType(typeof(decimal)).WithDefault(-1m);
        else builder.Column("Price").AsInt32();
        builder.Column("When").AsDateTime();
        builder.Column("Active").AsBoolean();
        builder.Column("Note").AsString();
        return Read(reader, new CsvReaderOptions {
            Schema = builder.Build(),
            MappingErrorValuePolicy = policy,
            TrimWhitespace = trim,
            Culture = CultureInfo.InvariantCulture
        });
    }

    private static string Read(string? reader, CsvReaderOptions options) {
        var result = new StringBuilder();
        try {
            // A null reader selects the UTF-8 byte stream; otherwise the decoded text overload.
            using CsvDataReader data = reader is null
                ? CsvDocument.OpenReader(new MemoryStream(Encoding.UTF8.GetBytes(Input)), options)
                : CsvDocument.OpenReader(new StringReader(reader), options);
            for (int column = 0; column < data.FieldCount; column++) {
                result.Append(data.GetName(column)).Append('=').Append(data.GetFieldType(column).Name).Append(' ');
            }
            result.Append('\n');
            while (true) {
                bool read;
                try {
                    read = data.Read();
                } catch (Exception exception) {
                    result.Append("read error ").Append(exception.GetType().Name).Append('\n');
                    break;
                }
                if (!read) break;
                for (int column = 0; column < data.FieldCount; column++) {
                    try {
                        object value = data.GetValue(column);
                        result.Append(value is DBNull ? "<null>" : Convert.ToString(value, CultureInfo.InvariantCulture))
                            .Append(':').Append(value.GetType().Name);
                    } catch (Exception exception) {
                        result.Append("error ").Append(exception.GetType().Name);
                    }
                    result.Append(" | ");
                }
                result.Append('\n');
            }
        } catch (Exception exception) {
            result.Append("open error ").Append(exception.GetType().Name);
        }
        return result.ToString();
    }
}
