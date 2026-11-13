using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>Plain text through OfficeIMO.Reader's text-and-metadata projection, split back into paragraphs.</summary>
internal static class WordReaderTextWorkload {
    private static readonly OfficeDocumentReader Reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

    internal static WordReadObservation Read(byte[] fixture) {
        OfficeDocumentReadResult result = Reader.ReadDocument(fixture, "fixture.docx", new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata,
            ComputeHashes = false
        });
        var observation = WordReadObservation.Empty;
        foreach (string line in result.GetText().Split('\n')) {
            string paragraph = line.TrimEnd('\r');
            if (paragraph.Length > 0) observation = observation.Add(paragraph);
        }
        return observation;
    }
}
