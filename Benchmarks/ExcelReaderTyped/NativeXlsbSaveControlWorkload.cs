#if OFFICEIMO_BENCHMARK_NEW_APIS
using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>The ordinary object construction and save shared by the warm and first-save controls.</summary>
    internal sealed class NativeXlsbSaveControlWorkload : IDisposable {
        internal const int Rows = 50_000;
        private const int OutputCapacity = 4 * 1024 * 1024;
        private List<TypedRecord> _records = [];
        private MemoryStream? _output;
        private bool _allText;
        private bool _shared;

        internal void Prepare(string shape, string storage) {
            BenchmarkInput.WriteDescription();
            if (shape is not ("Mixed" or "AllText") || storage is not ("Inline" or "Shared"))
                throw new ArgumentException("Native XLSB controls require Mixed/AllText and Inline/Shared.");
            _allText = shape == "AllText";
            _shared = storage == "Shared";
            _records = Enumerable.Range(1, Rows).Select(TypedWorkbookFixture.ExpectedRecord).ToList();
            _output = new MemoryStream(OutputCapacity);
            Console.WriteLine($"Native XLSB control input: rows={Rows}, shape={shape}, storage={storage}, "
                + $"capacity={OutputCapacity}; records and destination allocation excluded; "
                + "document creation, explicit InsertObjects, save and disposal included. "
                + "AllText invariant conversions occur in the insertion selectors. Input hash follows in cleanup.");
        }

        internal long Save() {
            MemoryStream output = _output ?? throw new InvalidOperationException("Prepare the native XLSB control first.");
            output.Position = 0;
            output.SetLength(0);
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                if (_allText)
                    sheet.InsertObjects(_records, ("Name", static record => record.Name),
                        ("Id", static record => record.Id.ToString(CultureInfo.InvariantCulture)),
                        ("Date", static record => record.Date.ToString("O", CultureInfo.InvariantCulture)),
                        ("Value", static record => record.Value.ToString("R", CultureInfo.InvariantCulture)));
                else
                    sheet.InsertObjects(_records, ("Name", static record => record.Name), ("Id", static record => record.Id),
                        ("Date", static record => record.Date), ("Value", static record => record.Value));
                document.Save(output, ExcelFileFormat.Xlsb, new ExcelSaveOptions { XlsbUseSharedStrings = _shared });
            }
            if (output.Capacity != OutputCapacity)
                throw new InvalidDataException("The native XLSB control destination grew; fixed-buffer evidence is invalid.");
            return output.Length;
        }

        internal void ValidateCurrentOutput(string lane) => NativeXlsbSaveControlQualification.Validate(
            (_output ?? throw new InvalidOperationException("Prepare the native XLSB control first.")).ToArray(),
            Rows, _allText, _shared, lane);

        // Called only by qualification and cleanup, after the measured save.
        internal static string[] TextFields(TypedRecord record) => [record.Name!, record.Id.ToString(CultureInfo.InvariantCulture),
            record.Date.ToString("O", CultureInfo.InvariantCulture), record.Value.ToString("R", CultureInfo.InvariantCulture)];

        private string InputHash() {
            using SHA256 hash = SHA256.Create();
            using (CryptoStream sink = new CryptoStream(Stream.Null, hash, CryptoStreamMode.Write, leaveOpen: true)) {
                using (Utf8JsonWriter writer = new Utf8JsonWriter(sink)) {
                    writer.WriteStartObject();
                    writer.WriteString("shape", _allText ? "AllText" : "Mixed");
                    writer.WritePropertyName("headers");
                    writer.WriteStartArray();
                    foreach (string header in TypedWorkbookFixture.Headers) writer.WriteStringValue(header);
                    writer.WriteEndArray();
                    writer.WritePropertyName("data");
                    writer.WriteStartArray();
                    foreach (TypedRecord record in _records) {
                        writer.WriteStartArray();
                        if (_allText) {
                            foreach (string text in TextFields(record)) writer.WriteStringValue(text);
                        } else {
                            writer.WriteStringValue(record.Name);
                            writer.WriteNumberValue(record.Id);
                            writer.WriteStringValue(record.Date.ToString("O", CultureInfo.InvariantCulture));
                            writer.WriteNumberValue(record.Value);
                        }
                        writer.WriteEndArray();
                    }
                    writer.WriteEndArray();
                    writer.WriteEndObject();
                }
                sink.FlushFinalBlock();
            }
            return Convert.ToHexString(hash.Hash!);
        }

        public void Dispose() {
            if (_output is null) return;
            _output.Dispose();
            _output = null;
            Console.WriteLine($"Native XLSB control input SHA256={InputHash()}.");
        }
    }
}
#endif
