using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>Plain text through OfficeIMO.Reader's text projections, split back into paragraphs.</summary>
internal static class WordReaderTextWorkload {
    private static readonly OfficeDocumentReader Reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

    /// <summary>The text-and-metadata projection, which builds the Word model.</summary>
    internal static WordReadObservation Read(byte[] fixture) => Read(fixture, ReaderProjection.TextAndMetadata);

    /// <summary>The text projection, which streams simple bodies without building the Word model.</summary>
    internal static WordReadObservation ReadStreamed(byte[] fixture) => Read(fixture, ReaderProjection.Text);

    private static WordReadObservation Read(byte[] fixture, ReaderProjection projection) {
        OfficeDocumentReadResult result = Reader.ReadDocument(fixture, "fixture.docx", new ReaderOptions {
            Projection = projection,
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
