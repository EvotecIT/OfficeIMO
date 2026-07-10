using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

public static partial class DocumentReaderOpenDocumentExtensions {
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
