using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

public static partial class DocumentReaderOpenDocumentExtensions {
    private static IEnumerable<ReaderChunk> ReadSpreadsheet(OdsDocument document, string sourceName, ReaderOptions options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        yield return new ReaderChunk {
            Id = BuildId(sourceName, "spreadsheet", 0),
            Kind = ReaderInputKind.OpenDocument,
            Location = new ReaderLocation { Path = sourceName, BlockIndex = 0, SourceBlockIndex = 0, SourceBlockKind = "spreadsheet" },
            Text = string.Empty,
            Warnings = new[] { "ODS extraction will be enabled with the sparse spreadsheet model." }
        };
    }

    private static IEnumerable<ReaderChunk> ReadPresentation(OdpPresentation document, string sourceName, ReaderOptions options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        yield return new ReaderChunk {
            Id = BuildId(sourceName, "presentation", 0),
            Kind = ReaderInputKind.OpenDocument,
            Location = new ReaderLocation { Path = sourceName, BlockIndex = 0, SourceBlockIndex = 0, SourceBlockKind = "presentation" },
            Text = string.Empty,
            Warnings = new[] { "ODP extraction will be enabled with the presentation model." }
        };
    }
}
