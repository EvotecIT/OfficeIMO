#nullable enable

using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using CsvHelper.Configuration;
using nietras.SeparatedValues;
using CsvHelperReader = CsvHelper.CsvReader;
using DataplatCsvDataReader = Dataplat.Dbatools.Csv.Reader.CsvDataReader;
using DataplatCsvReaderOptions = Dataplat.Dbatools.Csv.Reader.CsvReaderOptions;
using LumenWorksCsvReader = CsvReader.CsvReader;
using SepLib = nietras.SeparatedValues.Sep;
using SylvanCsvDataReader = Sylvan.Data.Csv.CsvDataReader;

namespace OfficeIMO.CSV.Benchmarks;

[MemoryDiagnoser]
[RankColumn]
[SimpleJob(RuntimeMoniker.Net80)]
public class CsvDbatoolsLibraryParityBenchmarks
{
    private const int FirstColumnIndex = 0;

    private string _smallCsvPath = string.Empty;
    private string _mediumCsvPath = string.Empty;
    private string _largeCsvPath = string.Empty;
    private string _wideCsvPath = string.Empty;
    private string _quotedCsvPath = string.Empty;
    private string _quickTestCsvPath = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        var dataDir = Path.Combine(AppContext.BaseDirectory, "TestData");
        Directory.CreateDirectory(dataDir);

        _smallCsvPath = Path.Combine(dataDir, "dbatools-small.csv");
        GenerateCsv(_smallCsvPath, 1_000, 10, quoteAll: false);

        _mediumCsvPath = Path.Combine(dataDir, "dbatools-medium.csv");
        GenerateCsv(_mediumCsvPath, 100_000, 10, quoteAll: false);

        _largeCsvPath = Path.Combine(dataDir, "dbatools-large.csv");
        GenerateCsv(_largeCsvPath, 1_000_000, 10, quoteAll: false);

        _wideCsvPath = Path.Combine(dataDir, "dbatools-wide.csv");
        GenerateCsv(_wideCsvPath, 100_000, 50, quoteAll: false);

        _quotedCsvPath = Path.Combine(dataDir, "dbatools-quoted.csv");
        GenerateCsv(_quotedCsvPath, 100_000, 10, quoteAll: true);

        _quickTestCsvPath = Path.Combine(dataDir, "dbatools-quicktest.csv");
        GenerateCsv(_quickTestCsvPath, 100_000, 10, quoteAll: false);
    }

    [Benchmark(Description = "OfficeIMO-Small")]
    [BenchmarkCategory("Small")]
    public int OfficeIMO_Small() => OfficeIMO_ReadFirstColumn(_smallCsvPath);

    [Benchmark(Baseline = true, Description = "Dataplat-Small")]
    [BenchmarkCategory("Small")]
    public int Dataplat_Small() => Dataplat_ReadFirstColumn(_smallCsvPath);

    [Benchmark(Description = "LumenWorks-Small")]
    [BenchmarkCategory("Small")]
    public int LumenWorks_Small() => LumenWorks_ReadFirstColumn(_smallCsvPath);

    [Benchmark(Description = "OfficeIMO-Medium")]
    [BenchmarkCategory("Medium")]
    public int OfficeIMO_Medium() => OfficeIMO_ReadFirstColumn(_mediumCsvPath);

    [Benchmark(Description = "Dataplat-Medium")]
    [BenchmarkCategory("Medium")]
    public int Dataplat_Medium() => Dataplat_ReadFirstColumn(_mediumCsvPath);

    [Benchmark(Description = "LumenWorks-Medium")]
    [BenchmarkCategory("Medium")]
    public int LumenWorks_Medium() => LumenWorks_ReadFirstColumn(_mediumCsvPath);

    [Benchmark(Description = "OfficeIMO-Large")]
    [BenchmarkCategory("Large")]
    public int OfficeIMO_Large() => OfficeIMO_ReadFirstColumn(_largeCsvPath);

    [Benchmark(Description = "Dataplat-Large")]
    [BenchmarkCategory("Large")]
    public int Dataplat_Large() => Dataplat_ReadFirstColumn(_largeCsvPath);

    [Benchmark(Description = "LumenWorks-Large")]
    [BenchmarkCategory("Large")]
    public int LumenWorks_Large() => LumenWorks_ReadFirstColumn(_largeCsvPath);

    [Benchmark(Description = "OfficeIMO-Wide")]
    [BenchmarkCategory("Wide")]
    public int OfficeIMO_Wide() => OfficeIMO_ReadFirstColumn(_wideCsvPath);

    [Benchmark(Description = "Dataplat-Wide")]
    [BenchmarkCategory("Wide")]
    public int Dataplat_Wide() => Dataplat_ReadFirstColumn(_wideCsvPath);

    [Benchmark(Description = "LumenWorks-Wide")]
    [BenchmarkCategory("Wide")]
    public int LumenWorks_Wide() => LumenWorks_ReadFirstColumn(_wideCsvPath);

    [Benchmark(Description = "OfficeIMO-Quoted")]
    [BenchmarkCategory("Quoted")]
    public int OfficeIMO_Quoted() => OfficeIMO_ReadFirstColumn(_quotedCsvPath);

    [Benchmark(Description = "Dataplat-Quoted")]
    [BenchmarkCategory("Quoted")]
    public int Dataplat_Quoted() => Dataplat_ReadFirstColumn(_quotedCsvPath);

    [Benchmark(Description = "LumenWorks-Quoted")]
    [BenchmarkCategory("Quoted")]
    public int LumenWorks_Quoted() => LumenWorks_ReadFirstColumn(_quotedCsvPath);

    [Benchmark(Description = "OfficeIMO-Medium-Modern")]
    [BenchmarkCategory("Modern")]
    public int OfficeIMO_Medium_Modern() => OfficeIMO_ReadFirstColumn(_mediumCsvPath);

    [Benchmark(Description = "Sep-Medium")]
    [BenchmarkCategory("Modern")]
    public int Sep_Medium() => Sep_ReadFirstColumn(_mediumCsvPath);

    [Benchmark(Description = "Sylvan-Medium")]
    [BenchmarkCategory("Modern")]
    public int Sylvan_Medium() => Sylvan_ReadFirstColumn(_mediumCsvPath);

    [Benchmark(Description = "CsvHelper-Medium")]
    [BenchmarkCategory("Modern")]
    public int CsvHelper_Medium() => CsvHelper_ReadFirstColumn(_mediumCsvPath);

    [Benchmark(Description = "Dataplat-Medium-Modern")]
    [BenchmarkCategory("Modern")]
    public int Dataplat_Medium_Modern() => Dataplat_ReadFirstColumn(_mediumCsvPath);

    [Benchmark(Description = "OfficeIMO-Large-Modern")]
    [BenchmarkCategory("ModernLarge")]
    public int OfficeIMO_Large_Modern() => OfficeIMO_ReadFirstColumn(_largeCsvPath);

    [Benchmark(Description = "Sep-Large")]
    [BenchmarkCategory("ModernLarge")]
    public int Sep_Large() => Sep_ReadFirstColumn(_largeCsvPath);

    [Benchmark(Description = "Sylvan-Large")]
    [BenchmarkCategory("ModernLarge")]
    public int Sylvan_Large() => Sylvan_ReadFirstColumn(_largeCsvPath);

    [Benchmark(Description = "CsvHelper-Large")]
    [BenchmarkCategory("ModernLarge")]
    public int CsvHelper_Large() => CsvHelper_ReadFirstColumn(_largeCsvPath);

    [Benchmark(Description = "Dataplat-Large-Modern")]
    [BenchmarkCategory("ModernLarge")]
    public int Dataplat_Large_Modern() => Dataplat_ReadFirstColumn(_largeCsvPath);

    [Benchmark(Description = "OfficeIMO-AllValues")]
    [BenchmarkCategory("AllValues")]
    public int OfficeIMO_AllValues() => OfficeIMO_ReadAllValues(_mediumCsvPath);

    [Benchmark(Description = "OfficeIMO-DataReader-AllValues")]
    [BenchmarkCategory("AllValues")]
    public int OfficeIMO_DataReader_AllValues() => OfficeIMO_DataReaderReadAllValues(_mediumCsvPath);

    [Benchmark(Description = "Dataplat-AllValues")]
    [BenchmarkCategory("AllValues")]
    public int Dataplat_AllValues()
    {
        var count = 0;
        var options = new DataplatCsvReaderOptions { BufferSize = 65_536 };
        using var reader = new DataplatCsvDataReader(_mediumCsvPath, options);
        var values = new object[reader.FieldCount];
        while (reader.Read())
        {
            count++;
            reader.GetValues(values);
        }

        return count;
    }

    [Benchmark(Description = "LumenWorks-AllValues")]
    [BenchmarkCategory("AllValues")]
    public int LumenWorks_AllValues()
    {
        var count = 0;
        using var textReader = new StreamReader(_mediumCsvPath);
        using var reader = new LumenWorksCsvReader(textReader, true);
        while (reader.ReadNextRecord())
        {
            count++;
            for (var i = 0; i < reader.FieldCount; i++)
            {
                _ = reader[i];
            }
        }

        return count;
    }

    [Benchmark(Description = "OfficeIMO-QuickTest-SingleColumn")]
    [BenchmarkCategory("QuickTest")]
    public int OfficeIMO_QuickTest_SingleColumn() => OfficeIMO_ReadFirstColumn(_quickTestCsvPath);

    [Benchmark(Description = "OfficeIMO-QuickTest-AllColumns")]
    [BenchmarkCategory("QuickTest")]
    public int OfficeIMO_QuickTest_AllColumns() => OfficeIMO_ReadAllValues(_quickTestCsvPath);

    [Benchmark(Description = "OfficeIMO-DataReader-QuickTest-SingleColumn")]
    [BenchmarkCategory("QuickTest")]
    public int OfficeIMO_DataReader_QuickTest_SingleColumn() => OfficeIMO_DataReaderReadFirstColumn(_quickTestCsvPath);

    [Benchmark(Description = "OfficeIMO-DataReader-QuickTest-AllColumns")]
    [BenchmarkCategory("QuickTest")]
    public int OfficeIMO_DataReader_QuickTest_AllColumns() => OfficeIMO_DataReaderReadAllValues(_quickTestCsvPath);

    [Benchmark(Description = "Sep-QuickTest-SingleColumn")]
    [BenchmarkCategory("QuickTest")]
    public int Sep_QuickTest_SingleColumn() => Sep_ReadFirstColumn(_quickTestCsvPath);

    [Benchmark(Description = "Sep-QuickTest-AllColumns")]
    [BenchmarkCategory("QuickTest")]
    public int Sep_QuickTest_AllColumns() => Sep_ReadAllValues(_quickTestCsvPath);

    [Benchmark(Description = "Sylvan-QuickTest-SingleColumn")]
    [BenchmarkCategory("QuickTest")]
    public int Sylvan_QuickTest_SingleColumn() => Sylvan_ReadFirstColumn(_quickTestCsvPath);

    [Benchmark(Description = "Sylvan-QuickTest-AllColumns")]
    [BenchmarkCategory("QuickTest")]
    public int Sylvan_QuickTest_AllColumns() => Sylvan_ReadAllValues(_quickTestCsvPath);

    [Benchmark(Description = "CsvHelper-QuickTest-SingleColumn")]
    [BenchmarkCategory("QuickTest")]
    public int CsvHelper_QuickTest_SingleColumn() => CsvHelper_ReadFirstColumn(_quickTestCsvPath);

    [Benchmark(Description = "CsvHelper-QuickTest-AllColumns")]
    [BenchmarkCategory("QuickTest")]
    public int CsvHelper_QuickTest_AllColumns() => CsvHelper_ReadAllValues(_quickTestCsvPath);

    [Benchmark(Description = "Dataplat-QuickTest-SingleColumn")]
    [BenchmarkCategory("QuickTest")]
    public int Dataplat_QuickTest_SingleColumn() => Dataplat_ReadFirstColumn(_quickTestCsvPath);

    [Benchmark(Description = "Dataplat-QuickTest-AllColumns")]
    [BenchmarkCategory("QuickTest")]
    public int Dataplat_QuickTest_AllColumns() => Dataplat_ReadAllValues(_quickTestCsvPath);

    [Benchmark(Description = "LumenWorks-QuickTest-SingleColumn")]
    [BenchmarkCategory("QuickTest")]
    public int LumenWorks_QuickTest_SingleColumn() => LumenWorks_ReadFirstColumn(_quickTestCsvPath);

    [Benchmark(Description = "LumenWorks-QuickTest-AllColumns")]
    [BenchmarkCategory("QuickTest")]
    public int LumenWorks_QuickTest_AllColumns() => LumenWorks_ReadAllValues(_quickTestCsvPath);

    private static int OfficeIMO_ReadFirstColumn(string path)
    {
        var visitor = new OfficeImoFirstColumnVisitor();
        CsvDocument.ReadFieldSpans(path, ref visitor, new CsvLoadOptions { SkipInitialRecords = 1, DetectDelimiter = false });
        if (visitor.FirstColumnLength < 0)
        {
            throw new InvalidOperationException("Unexpected negative field length.");
        }

        return visitor.RowCount;
    }

    private static int OfficeIMO_ReadAllValues(string path)
    {
        var visitor = new OfficeImoAllValuesVisitor();
        CsvDocument.ReadFieldSpans(path, ref visitor, new CsvLoadOptions { SkipInitialRecords = 1, DetectDelimiter = false });
        if (visitor.FieldLength < 0)
        {
            throw new InvalidOperationException("Unexpected negative field length.");
        }

        return visitor.RowCount;
    }

    private static int OfficeIMO_DataReaderReadFirstColumn(string path)
    {
        var count = 0;
        var document = CsvDocument.Load(path, new CsvLoadOptions { Mode = CsvLoadMode.Stream, DetectDelimiter = false });
        using var reader = document.CreateDataReader();
        while (reader.Read())
        {
            count++;
            _ = reader.GetValue(FirstColumnIndex);
        }

        return count;
    }

    private static int OfficeIMO_DataReaderReadAllValues(string path)
    {
        var count = 0;
        var document = CsvDocument.Load(path, new CsvLoadOptions { Mode = CsvLoadMode.Stream, DetectDelimiter = false });
        using var reader = document.CreateDataReader();
        var values = new object[reader.FieldCount];
        while (reader.Read())
        {
            count++;
            reader.GetValues(values);
        }

        return count;
    }

    private static int Dataplat_ReadFirstColumn(string path)
    {
        var count = 0;
        using var reader = new DataplatCsvDataReader(path);
        while (reader.Read())
        {
            count++;
            _ = reader.GetValue(FirstColumnIndex);
        }

        return count;
    }

    private static int Dataplat_ReadAllValues(string path)
    {
        var count = 0;
        using var reader = new DataplatCsvDataReader(path);
        while (reader.Read())
        {
            count++;
            for (var i = 0; i < reader.FieldCount; i++)
            {
                _ = reader.GetValue(i);
            }
        }

        return count;
    }

    private static int LumenWorks_ReadFirstColumn(string path)
    {
        var count = 0;
        using var textReader = new StreamReader(path);
        using var reader = new LumenWorksCsvReader(textReader, true);
        while (reader.ReadNextRecord())
        {
            count++;
            _ = reader[FirstColumnIndex];
        }

        return count;
    }

    private static int LumenWorks_ReadAllValues(string path)
    {
        var count = 0;
        using var textReader = new StreamReader(path);
        using var reader = new LumenWorksCsvReader(textReader, true);
        while (reader.ReadNextRecord())
        {
            count++;
            for (var i = 0; i < reader.FieldCount; i++)
            {
                _ = reader[i];
            }
        }

        return count;
    }

    private static int Sep_ReadFirstColumn(string path)
    {
        var count = 0;
        using var reader = SepLib.Reader().FromFile(path);
        foreach (var row in reader)
        {
            count++;
            _ = row[FirstColumnIndex].ToString();
        }

        return count;
    }

    private static int Sep_ReadAllValues(string path)
    {
        var count = 0;
        using var reader = SepLib.Reader().FromFile(path);
        foreach (var row in reader)
        {
            count++;
            for (var i = 0; i < row.ColCount; i++)
            {
                _ = row[i].ToString();
            }
        }

        return count;
    }

    private static int Sylvan_ReadFirstColumn(string path)
    {
        var count = 0;
        using var textReader = new StreamReader(path);
        using var reader = SylvanCsvDataReader.Create(textReader);
        while (reader.Read())
        {
            count++;
            _ = reader.GetString(FirstColumnIndex);
        }

        return count;
    }

    private static int Sylvan_ReadAllValues(string path)
    {
        var count = 0;
        using var textReader = new StreamReader(path);
        using var reader = SylvanCsvDataReader.Create(textReader);
        while (reader.Read())
        {
            count++;
            for (var i = 0; i < reader.FieldCount; i++)
            {
                _ = reader.GetString(i);
            }
        }

        return count;
    }

    private static int CsvHelper_ReadFirstColumn(string path)
    {
        var count = 0;
        var config = new CsvConfiguration(CultureInfo.InvariantCulture) { HasHeaderRecord = true };
        using var textReader = new StreamReader(path);
        using var csv = new CsvHelperReader(textReader, config);
        csv.Read();
        csv.ReadHeader();
        while (csv.Read())
        {
            count++;
            _ = csv.GetField(FirstColumnIndex);
        }

        return count;
    }

    private static int CsvHelper_ReadAllValues(string path)
    {
        var count = 0;
        var config = new CsvConfiguration(CultureInfo.InvariantCulture) { HasHeaderRecord = true };
        using var textReader = new StreamReader(path);
        using var csv = new CsvHelperReader(textReader, config);
        csv.Read();
        csv.ReadHeader();
        while (csv.Read())
        {
            count++;
            for (var i = 0; i < csv.Parser.Count; i++)
            {
                _ = csv.GetField(i);
            }
        }

        return count;
    }

    private static void GenerateCsv(string path, int rows, int columns, bool quoteAll)
    {
        if (File.Exists(path))
        {
            return;
        }

        using var writer = new StreamWriter(path, append: false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.WriteLine(string.Join(",", Enumerable.Range(0, columns).Select(static i => string.Create(CultureInfo.InvariantCulture, $"Column{i}"))));

        var random = new Random(42);
        var builder = new StringBuilder();
        for (var row = 0; row < rows; row++)
        {
            builder.Clear();
            for (var column = 0; column < columns; column++)
            {
                if (column > 0)
                {
                    builder.Append(',');
                }

                var value = column switch
                {
                    0 => row.ToString(CultureInfo.InvariantCulture),
                    1 => string.Create(CultureInfo.InvariantCulture, $"Name{row}"),
                    2 => random.Next(1, 100).ToString(CultureInfo.InvariantCulture),
                    3 => random.NextDouble().ToString("F4", CultureInfo.InvariantCulture),
                    4 => DateTime.Now.AddDays(-random.Next(365)).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                    5 => random.Next(0, 2) == 0 ? "true" : "false",
                    _ => string.Create(CultureInfo.InvariantCulture, $"Value{row}_{column}")
                };

                if (quoteAll)
                {
                    builder.Append('"');
                    builder.Append(value);
                    builder.Append('"');
                }
                else
                {
                    builder.Append(value);
                }
            }

            writer.WriteLine(builder.ToString());
        }
    }

    private struct OfficeImoFirstColumnVisitor : ICsvProjectedFieldSpanVisitor
    {
        public int RowCount { get; private set; }

        public int FirstColumnLength { get; private set; }

        public bool ShouldVisitField(int recordIndex, int fieldIndex) => fieldIndex == FirstColumnIndex;

        public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
        {
            RowCount++;
            FirstColumnLength += value.Length;
        }

        public bool TryVisitEscapedField(int recordIndex, int fieldIndex, ReadOnlySpan<char> escapedValue, int unescapedLength)
        {
            RowCount++;
            FirstColumnLength += unescapedLength;
            return true;
        }
    }

    private struct OfficeImoAllValuesVisitor : ICsvFieldSpanVisitor
    {
        public int RowCount { get; private set; }

        public int FieldLength { get; private set; }

        public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
        {
            if (fieldIndex == FirstColumnIndex)
            {
                RowCount++;
            }

            FieldLength += value.Length;
        }

        public bool TryVisitEscapedField(int recordIndex, int fieldIndex, ReadOnlySpan<char> escapedValue, int unescapedLength)
        {
            if (fieldIndex == FirstColumnIndex)
            {
                RowCount++;
            }

            FieldLength += unescapedLength;
            return true;
        }
    }
}
