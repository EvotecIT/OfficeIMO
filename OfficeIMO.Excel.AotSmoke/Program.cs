using System.Collections;
using System.Collections.Generic;
using System.Data;
using OfficeIMO.Data;
using OfficeIMO.Excel;

string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-AotSmoke-" + Guid.NewGuid().ToString("N") + ".xlsx");
try {
    using (var asyncRows = new MemoryStream()) {
        ExcelDataSetImportResult result = await ExcelDocument.WriteRowsAsync(
            asyncRows,
            CreateRowsAsync(),
            ["Region", "Revenue"],
            static (writer, row) => writer.Write(row.Region).Write(row.Revenue));
        if (result.Range != "A1:B3" || result.RowCount != 2) {
            throw new InvalidOperationException("The asynchronous row writer returned an unexpected range.");
        }
    }

    using (ExcelDocument document = ExcelDocument.Create(path)) {
        var sales = new DataTable("Sales");
        sales.Columns.Add("Region", typeof(string));
        sales.Columns.Add("Revenue", typeof(decimal));
        sales.Columns.Add("Date", typeof(DateOnly));
        sales.Columns.Add("Time", typeof(TimeOnly));
        sales.Rows.Add("North", 1250000M, new DateOnly(2026, 8, 6), new TimeOnly(14, 35, 12));
        sales.Rows.Add("South", 980000M, new DateOnly(2026, 8, 7), new TimeOnly(9, 15, 0));

        using (var dataReaderPackage = new MemoryStream()) {
            using var dataReader = sales.CreateDataReader();
            ExcelDataSetImportResult dataReaderResult = ExcelDocument.WriteDataReader(
                dataReaderPackage,
                dataReader,
                new ExcelTabularWriteOptions {
                    IncludeCellReferences = false,
                    UseSharedStrings = false
                });
            if (dataReaderResult.Range != "A1:D3" || dataReaderResult.RowCount != 2) {
                throw new InvalidOperationException("The AOT-safe DataReader writer returned an unexpected range.");
            }

            dataReaderPackage.Position = 0;
            using ExcelWorkbookDataReader dataReaderWorkbook = ExcelDocument.OpenDataReader(dataReaderPackage);
            if (!dataReaderWorkbook.Read() || dataReaderWorkbook.GetString(0) != "North" ||
                !dataReaderWorkbook.Read() ||
                Convert.ToDecimal(dataReaderWorkbook.GetValue(1), System.Globalization.CultureInfo.InvariantCulture) != 980000M) {
                throw new InvalidOperationException("The AOT-safe DataReader writer lost typed values.");
            }
        }

        ExcelSheet sheet = document.AddWorksheet("NativeAOT data");
        string range = sheet.InsertDataTableAsTable(sales, tableName: "Sales");
        if (range != "A1:D3") {
            throw new InvalidOperationException($"The Excel table used the unexpected range '{range}'.");
        }

        var dictionaryRows = new[] {
            new GenericOnlyDictionaryRow<int>(("Score", 10), ("Rank", 1)),
            new GenericOnlyDictionaryRow<int>(("Score", 20), ("Rank", 2))
        };
        document.AsFluent()
            .Sheet("Dictionary builder", builder => builder.RowsFrom(dictionaryRows))
            .End();
        document.Compose("Dictionary composer", composer =>
            composer.TableFrom(dictionaryRows, title: "Scores"));
        document.Save();
    }

    using ExcelDocument reopened = ExcelDocument.Load(path);
    if (reopened.Sheets.Count != 3 || reopened.Sheets[0].Name != "NativeAOT data") {
        throw new InvalidOperationException("The Excel round trip lost its worksheet.");
    }
    if (!reopened.Sheets[0].TryGetCellText(2, 1, out string region) || region != "North") {
        throw new InvalidOperationException("The Excel round trip lost its typed table data.");
    }
    AotSalesRow mappedRow = reopened.Sheets[0].RowsAs<AotSalesRow>(map => map
        .FromColumn<string>("Region", static (row, value) => { row.Region = value; return row; })
        .FromColumn<decimal>("Revenue", static (row, value) => { row.Revenue = value; return row; })
        .FromColumn<DateOnly>("Date", static (row, value) => { row.Date = value; return row; })
        .FromColumn<TimeOnly>("Time", static (row, value) => { row.Time = value; return row; }))
        .First();
    if (mappedRow.Region != "North" || mappedRow.Revenue != 1250000M ||
        mappedRow.Date != new DateOnly(2026, 8, 6) || mappedRow.Time != new TimeOnly(14, 35, 12)) {
        throw new InvalidOperationException("The AOT-safe typed-row mapping returned unexpected data.");
    }
    if (!reopened["Dictionary builder"].TryGetCellText(1, 1, out string builderHeader) || builderHeader != "Score" ||
        !reopened["Dictionary builder"].TryGetCellText(2, 1, out string builderValue) || builderValue != "10") {
        throw new InvalidOperationException("RowsFrom did not preserve generic-only dictionary columns under NativeAOT.");
    }
    if (!reopened["Dictionary composer"].TryGetCellText(2, 1, out string composerRankHeader) || composerRankHeader != "Rank" ||
        !reopened["Dictionary composer"].TryGetCellText(2, 2, out string composerScoreHeader) || composerScoreHeader != "Score" ||
        !reopened["Dictionary composer"].TryGetCellText(3, 1, out string composerRankValue) || composerRankValue != "1" ||
        !reopened["Dictionary composer"].TryGetCellText(3, 2, out string composerScoreValue) || composerScoreValue != "10") {
        throw new InvalidOperationException("TableFrom did not preserve generic-only dictionary columns under NativeAOT.");
    }

    Console.WriteLine("PASS | Excel DataReader, typed, and generic-only dictionary tables create, save, and reload");
} finally {
    if (File.Exists(path)) File.Delete(path);
}

static async IAsyncEnumerable<SalesRow> CreateRowsAsync() {
    await Task.CompletedTask;
    yield return new SalesRow("North", 1250000M);
    yield return new SalesRow("South", 980000M);
}

internal readonly record struct SalesRow(string Region, decimal Revenue);

internal sealed class AotSalesRow {
    public string Region { get; set; } = string.Empty;
    public decimal Revenue { get; set; }
    public DateOnly Date { get; set; }
    public TimeOnly Time { get; set; }
}

file sealed class GenericOnlyDictionaryRow<TValue> : IReadOnlyDictionary<string, TValue> {
    private readonly KeyValuePair<string, TValue>[] _entries;
    private readonly Dictionary<string, TValue> _lookup;

    internal GenericOnlyDictionaryRow(params (string Key, TValue Value)[] entries) {
        _entries = entries.Select(entry => new KeyValuePair<string, TValue>(entry.Key, entry.Value)).ToArray();
        _lookup = _entries.ToDictionary(entry => entry.Key, entry => entry.Value, StringComparer.OrdinalIgnoreCase);
    }

    public TValue this[string key] => _lookup[key];
    public IEnumerable<string> Keys => _entries.Select(entry => entry.Key);
    public IEnumerable<TValue> Values => _entries.Select(entry => entry.Value);
    public int Count => _entries.Length;
    public bool ContainsKey(string key) => _lookup.ContainsKey(key);
    public bool TryGetValue(string key, out TValue value) => _lookup.TryGetValue(key, out value!);
    public IEnumerator<KeyValuePair<string, TValue>> GetEnumerator() => ((IEnumerable<KeyValuePair<string, TValue>>)_entries).GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
