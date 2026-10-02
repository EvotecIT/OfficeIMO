using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;
using System.Text;

const string csv = "Name,Score\nAlice,42\nBob,51\n";
using MemoryStream stream = new(Encoding.UTF8.GetBytes(csv), writable: false);
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddCsvHandler(new CsvReadOptions {
        ChunkRows = 1,
        IncludeMarkdown = true
    })
    .Build();
List<ReaderChunk> chunks = reader.Read(stream, "scores.csv").ToList();

if (chunks.Count == 0 || chunks.Any(chunk => chunk.Kind != ReaderInputKind.Csv)) {
    throw new InvalidOperationException("Reader did not emit normalized CSV chunks.");
}
if (!chunks.Any(chunk => chunk.Tables is { Count: > 0 })) {
    throw new InvalidOperationException("Reader did not emit structured table data.");
}

Console.WriteLine("PASS | Reader CSV normalized extraction");

var textReader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
using var textStream = new MemoryStream(Encoding.UTF8.GetBytes("Unicode 漢字 🙂"));
var child = textReader.ReadDocument(textStream, "child.txt", ReaderOptions.CreateSafeIngestion());
var container = new OfficeDocumentReadResult {
    NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "container::child.txt", Document = child } }
};
var restored = OfficeDocumentReadResultJson.Deserialize(container.ToJson());
if (restored.NestedDocuments.Count != 1 || restored.NestedDocuments[0].Document.Chunks[0].Text != "Unicode 漢字 🙂")
    throw new InvalidOperationException("NativeAOT nested result transport lost content.");
if (!textReader.GetCapabilityManifestJson().Contains("formatQualifications", StringComparison.Ordinal))
    throw new InvalidOperationException("NativeAOT capability transport lost format qualification.");
using var incremental = new MemoryStream(Encoding.UTF8.GetBytes(new string('x', 600)));
if (textReader.EnumerateChunks(incremental, "large.txt", new ReaderOptions { ComputeHashes = false, MaxChars = 256 }).Count() != 3)
    throw new InvalidOperationException("NativeAOT incremental text extraction lost chunks.");
Console.WriteLine("PASS | Reader incremental text and nested transport");
