using System.Text;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddAllOfficeIMOHandlers()
    .Build();

ReaderHandlerCapability[] capabilities = reader.GetCapabilities()
    .Where(static capability => capability.Origin == ReaderHandlerOrigin.OfficeIMO)
    .ToArray();

string[] requiredHandlers = {
    "officeimo.reader.asciidoc",
    "officeimo.reader.csv",
    "officeimo.reader.email",
    "officeimo.reader.epub",
    "officeimo.reader.excel",
    "officeimo.reader.html",
    "officeimo.reader.image",
    "officeimo.reader.json",
    "officeimo.reader.latex",
    "officeimo.reader.markdown",
    "officeimo.reader.notebook",
    "officeimo.reader.onenote",
    "officeimo.reader.opendocument",
    "officeimo.reader.pdf",
    "officeimo.reader.powerpoint",
    "officeimo.reader.rtf",
    "officeimo.reader.subtitles",
    "officeimo.reader.visio",
    "officeimo.reader.word",
    "officeimo.reader.xml",
    "officeimo.reader.yaml",
    "officeimo.reader.zip"
};

foreach (string handlerId in requiredHandlers) {
    if (!capabilities.Any(capability => capability.Id == handlerId)) {
        throw new InvalidOperationException($"The NativeAOT reader preset did not register '{handlerId}'.");
    }
}

const string csv = "Name,Score\nAlice,42\nBob,51\n";
using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv), writable: false);
List<ReaderChunk> chunks = reader.Read(stream, "scores.csv").ToList();
if (!chunks.Any(chunk => chunk.Kind == ReaderInputKind.Csv && chunk.Tables is { Count: > 0 })) {
    throw new InvalidOperationException("The NativeAOT all-formats reader lost structured CSV extraction.");
}

Console.WriteLine($"PASS | Reader all-formats preset registered {capabilities.Length} in-process handlers and extracted CSV");
