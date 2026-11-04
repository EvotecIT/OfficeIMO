using BenchmarkDotNet.Attributes;

namespace OfficeIMO.CSV.Benchmarks;

public partial class CsvBenchmarks
{
    /// <summary>Reads the prepared UTF-8 input through the incremental stream reader.</summary>
    [Benchmark]
    public int OfficeIMO_ReadStreamGetStrings()
    {
        using var stream = new MemoryStream(_csvUtf8, writable: false);
        using var csv = CsvDocument.OpenReader(stream);
        var checksum = 0;
        while (csv.Read())
        {
            for (var column = 0; column < csv.FieldCount; column++)
            {
                string value = csv.GetString(column);
                checksum += 1 + value.Length;
                if (_validateRawRead) ValidateRawField(column, value);
            }
            if (_validateRawRead) CompleteRawRow();
        }
        return checksum;
    }
}
