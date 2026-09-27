using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using PowerForge;

// Workload only: PowerForge owns measurement ordering, timing, sampling and artifacts.
public sealed class CsvSustainedReadWorkload
{
    private readonly string _path;
    private readonly int _rows;
    private readonly bool _multiline;
    private readonly int _degree;
    private readonly int _batch;
    private readonly Dictionary<string, long> _expectedByOperation = new Dictionary<string, long>();
    private long _expected;
    public BenchmarkMemoryObservation Memory { get; private set; }
    public long Checksum { get; private set; }
    public long InputBytes => new FileInfo(_path).Length;

    public CsvSustainedReadWorkload(string path, int rows, bool multiline, int degree, int batch)
    {
        _path = path; _rows = rows; _multiline = multiline; _degree = degree; _batch = batch;
        using (var writer = new StreamWriter(path, false, new UTF8Encoding(false)))
        {
            writer.NewLine = "\r\n";
            writer.WriteLine("Id,Name,Notes,Score,Created,Enabled");
            for (int id = 1; id <= rows; id++)
            {
                Row row = Expected(id);
                writer.Write(id.ToString(CultureInfo.InvariantCulture)); writer.Write(',');
                writer.Write(row.Name); writer.Write(",\"");
                writer.Write(row.Notes.Replace("\"", "\"\"")); writer.Write("\",");
                writer.Write(row.Score.ToString(CultureInfo.InvariantCulture)); writer.Write(',');
                writer.Write(row.Created.ToString("O", CultureInfo.InvariantCulture)); writer.Write(',');
                writer.WriteLine(row.Enabled ? "true" : "false");
            }
        }
    }

    public void Prepare(string operation)
    {
        _expectedByOperation[operation] = ExpectedChecksum(operation);
        // Validate every field and source order outside timed operations, independently of checksums.
        RunAsync(false, operation, true).GetAwaiter().GetResult();
        RunAsync(true, operation, true).GetAwaiter().GetResult();
    }

    public void Execute(bool incremental, string operation, bool sampleMemory)
    {
        Memory = null;
        _expected = _expectedByOperation[operation];
        if (!sampleMemory) { Checksum = RunAsync(incremental, operation, false).GetAwaiter().GetResult(); return; }
        using (var probe = BenchmarkMemoryProbe.Start())
        {
            Checksum = RunAsync(incremental, operation, false).GetAwaiter().GetResult();
            Memory = probe.Complete();
        }
    }

    public void Validate()
    {
        if (Checksum != _expected) throw new InvalidDataException("CSV sustained-read checksum mismatch.");
    }

    private async Task<long> RunAsync(bool incremental, string operation, bool validate)
    {
        var options = new CsvLoadOptions { Culture = CultureInfo.InvariantCulture };
        using (DbDataReader reader = incremental
            ? await CsvDocument.OpenStreamingDataReaderAsync(_path, options).ConfigureAwait(false)
            : await CsvDocument.OpenDataReaderAsync(_path, options).ConfigureAwait(false))
        {
            if (operation == "FirstRow")
            {
                if (!await reader.ReadAsync().ConfigureAwait(false)) throw new InvalidDataException("Missing first row.");
                long hash = 17;
                string[] expected = Fields(Expected(1));
                for (int column = 0; column < expected.Length; column++)
                {
                    string value = reader.GetString(column);
                    if (validate && value != expected[column]) throw new InvalidDataException("First-row field mismatch.");
                    hash = Hash(hash, value);
                }
                return hash;
            }
            if (operation == "AllRowsAsync")
            {
                int rows = 0;
                long hash = 17;
                while (await reader.ReadAsync().ConfigureAwait(false))
                {
                    rows++;
                    string[] expected = validate ? Fields(Expected(rows)) : null;
                    for (int column = 0; column < 6; column++)
                    {
                        string value = reader.GetString(column);
                        if (validate && value != expected[column])
                            throw new InvalidDataException("Async CSV field/order mismatch at row " + rows);
                        hash = Hash(hash, value);
                    }
                }
                if (rows != _rows) throw new InvalidDataException("Async CSV row-count mismatch.");
                return hash;
            }
            IEnumerable<Row> projected = operation == "TypedParallel"
                ? reader.RowsAsParallel<Row>(new ParallelRowMappingOptions { MaxDegreeOfParallelism = _degree, BatchSize = _batch })
                : operation == "TypedSequential" ? reader.RowsAs<Row>()
                : throw new ArgumentException("Unknown CSV operation.", nameof(operation));
            int count = 0;
            long checksum = 17;
            foreach (Row row in projected)
            {
                count++;
                if (validate)
                {
                    Row expected = Expected(count);
                    if (row.Id != expected.Id || row.Name != expected.Name || row.Notes != expected.Notes ||
                        row.Score != expected.Score || row.Created != expected.Created || row.Enabled != expected.Enabled)
                        throw new InvalidDataException("Typed CSV field/order mismatch at row " + count);
                }
                checksum = HashRow(checksum, row);
            }
            if (count != _rows) throw new InvalidDataException("CSV row-count mismatch.");
            return checksum;
        }
    }

    private long ExpectedChecksum(string operation)
    {
        long hash = 17;
        if (operation == "FirstRow")
        {
            foreach (string value in Fields(Expected(1))) hash = Hash(hash, value);
            return hash;
        }
        if (operation == "AllRowsAsync")
        {
            for (int id = 1; id <= _rows; id++)
                foreach (string value in Fields(Expected(id))) hash = Hash(hash, value);
            return hash;
        }
        for (int id = 1; id <= _rows; id++) hash = HashRow(hash, Expected(id));
        return hash;
    }

    private Row Expected(int id) => new Row {
        Id = id, Name = "Person " + id.ToString(CultureInfo.InvariantCulture),
        Notes = "Unique note " + id.ToString(CultureInfo.InvariantCulture) + (_multiline ? "\r\nsecond line, \"quoted\" Zażółć" : ""),
        Score = id / 100m, Created = new DateTime(2024, 1, 1).AddSeconds(id), Enabled = id % 2 == 0
    };

    private static string[] Fields(Row row) => new[] { row.Id.ToString(CultureInfo.InvariantCulture), row.Name,
        row.Notes, row.Score.ToString(CultureInfo.InvariantCulture), row.Created.ToString("O", CultureInfo.InvariantCulture),
        row.Enabled ? "true" : "false" };

    private static long HashRow(long hash, Row row)
    {
        unchecked
        {
            hash = hash * 31 + row.Id;
            hash = Hash(Hash(hash, row.Name), row.Notes);
            hash = hash * 31 + (long)(row.Score * 100);
            hash = hash * 31 + row.Created.Ticks;
            return hash * 31 + (row.Enabled ? 1 : 0);
        }
    }
    private static long Hash(long hash, string value)
    {
        unchecked { foreach (char character in value) hash = hash * 31 + character; return hash * 31 + value.Length; }
    }

    public sealed class Row
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Notes { get; set; }
        public decimal Score { get; set; }
        public DateTime Created { get; set; }
        public bool Enabled { get; set; }
    }
}
