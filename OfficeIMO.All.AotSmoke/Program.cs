using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeIMO.Adf;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using OfficeIMO.Excel;
using OfficeIMO.GoogleWorkspace.Auth.GoogleApis;
using OfficeIMO.Reader;

var adfAttributes = new ReadOnlyObjectDictionary(new Dictionary<string, object?> {
    ["html"] = "<strong>Ready</strong>",
    ["enabled"] = true
});
JsonElement adfValue = new AdfNode("extension")
    .SetAttribute("parameters", adfAttributes)
    .Attributes["parameters"];
if (adfValue.ValueKind != JsonValueKind.Object || !adfValue.GetProperty("enabled").GetBoolean()) {
    throw new InvalidOperationException("The ADF read-only dictionary did not retain its JSON object shape.");
}

string readerJson = OfficeDocumentReadResultJson.Serialize(new OfficeDocumentReadResult {
    Metadata = new[] {
        new OfficeDocumentMetadataEntry {
            Id = "metadata-1",
            Category = "core",
            Name = "fixture",
            Attributes = new Dictionary<string, string> {
                ["zeta"] = "last",
                ["alpha"] = "first"
            }
        }
    }
});
int alphaIndex = readerJson.IndexOf("\"alpha\"", StringComparison.Ordinal);
int zetaIndex = readerJson.IndexOf("\"zeta\"", StringComparison.Ordinal);
if (alphaIndex < 0 || zetaIndex < 0 || alphaIndex > zetaIndex) {
    throw new InvalidOperationException("Reader metadata attributes were not serialized deterministically.");
}

var store = new InMemoryTokenStore();
var adapter = new GoogleApisDataStoreAdapter(store);
await adapter.StoreAsync("officeimo-aot", "token-marker");
var value = await adapter.GetAsync<string>("officeimo-aot");

if (!string.Equals(value, "token-marker", StringComparison.Ordinal)) {
    throw new InvalidOperationException("The Google APIs data-store adapter did not round-trip its value.");
}

using var tabularStream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("Id,Name\n1,Ada\n"));
using var csvReader = CsvDocument.OpenDataReader(
    tabularStream,
    new CsvLoadOptions(),
    new CsvDataReaderOptions { InferSchema = true });
if (!csvReader.Read() || csvReader.GetInt32(0) != 1 || csvReader.GetString(1) != "Ada") {
    throw new InvalidOperationException("The canonical tabular reader did not read its NativeAOT CSV fixture.");
}

using var parallelReader = CsvDocument.OpenTextDataReader("Id,Name\n1,Ada\n2,Grace\n");
AotCsvRow[] parallelRows = parallelReader.RowsAsParallel<AotCsvRow>(map => map
    .FromColumn<int>("Id", static (row, id) => { row.Id = id; return row; })
    .FromColumn<string>("Name", static (row, name) => { row.Name = name; return row; }),
    new ParallelRowMappingOptions {
        MaxDegreeOfParallelism = 2,
        BatchSize = 1
    }).ToArray();
if (parallelRows.Length != 2
    || parallelRows[0].Id != 1
    || parallelRows[0].Name != "Ada"
    || parallelRows[1].Id != 2
    || parallelRows[1].Name != "Grace") {
    throw new InvalidOperationException("The explicit ordered-parallel CSV mapping did not survive NativeAOT.");
}

using var excelSourceStream = new System.IO.MemoryStream();
ExcelDocument.WriteRows(
    excelSourceStream,
    new[] {
        new AotExcelRow { Id = 1, Name = "Ada" },
        new AotExcelRow { Id = 2, Name = "Grace" }
    },
    new[] { "Id", "Name" },
    static (writer, row) => writer.Write(row.Id).Write(row.Name),
    new ExcelTabularWriteOptions {
        IncludeCellReferences = false,
        UseSharedStrings = false
    });
excelSourceStream.Position = 0;
using var excelSourceReader = ExcelDocument.OpenDataReader(
    excelSourceStream,
    new ExcelReadOptions {
        SheetName = "Data",
        InferSchema = true
    });
using var excelStream = new System.IO.MemoryStream();
ExcelDocument.WriteDataReader(
    excelStream,
    excelSourceReader,
    new ExcelTabularWriteOptions {
        IncludeCellReferences = false,
        UseSharedStrings = false
    });
excelStream.Position = 0;
using var excelReader = ExcelDocument.OpenDataReader(
    excelStream,
    new ExcelReadOptions {
        SheetName = "Data",
        InferSchema = true
    });
AotExcelRow[] excelRows = excelReader.RowsAsParallel(
    static record => new AotExcelRow {
        Id = record.GetInt32(0),
        Name = record.GetString(1)
    },
    new ParallelRowMappingOptions {
        MaxDegreeOfParallelism = 2,
        BatchSize = 1
    }).ToArray();
if (excelRows.Length != 2
    || excelRows[0].Id != 1
    || excelRows[0].Name != "Ada"
    || excelRows[1].Id != 2
    || excelRows[1].Name != "Grace") {
    throw new InvalidOperationException("The IDataReader XLSX write fallback and ordered-parallel Excel mapping did not survive NativeAOT.");
}

Console.WriteLine("PASS | production libraries fully rooted; Google APIs token-store plus CSV and Excel read/write and ordered-parallel contracts passed from NativeAOT.");

file sealed class AotCsvRow {
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
}

file sealed class AotExcelRow {
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
}

file sealed class InMemoryTokenStore : IGoogleWorkspaceTokenStore {
    private readonly Dictionary<string, object?> _values = new(StringComparer.Ordinal);

    public Task StoreAsync<T>(string key, T value) {
        _values[key] = value;
        return Task.CompletedTask;
    }

    public Task DeleteAsync<T>(string key) {
        _values.Remove(key);
        return Task.CompletedTask;
    }

    public Task<T?> GetAsync<T>(string key) {
        return Task.FromResult(_values.TryGetValue(key, out var value) ? (T?)value : default);
    }

    public Task ClearAsync() {
        _values.Clear();
        return Task.CompletedTask;
    }
}

file sealed class ReadOnlyObjectDictionary : IReadOnlyDictionary<string, object?> {
    private readonly IReadOnlyDictionary<string, object?> _values;

    public ReadOnlyObjectDictionary(IReadOnlyDictionary<string, object?> values) {
        _values = values;
    }

    public object? this[string key] => _values[key];
    public IEnumerable<string> Keys => _values.Keys;
    public IEnumerable<object?> Values => _values.Values;
    public int Count => _values.Count;
    public bool ContainsKey(string key) => _values.ContainsKey(key);
    public bool TryGetValue(string key, out object? value) => _values.TryGetValue(key, out value);
    public IEnumerator<KeyValuePair<string, object?>> GetEnumerator() => _values.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
