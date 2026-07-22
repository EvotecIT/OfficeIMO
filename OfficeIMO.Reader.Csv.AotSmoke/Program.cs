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
