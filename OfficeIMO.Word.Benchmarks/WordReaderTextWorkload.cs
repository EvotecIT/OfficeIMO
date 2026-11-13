using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>Plain text through OfficeIMO.Reader's text-and-metadata projection, split back into paragraphs.</summary>
internal static class WordReaderTextWorkload {
    private static readonly OfficeDocumentReader Reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();
    private static readonly OfficeDocumentReader StreamingReader = new OfficeDocumentReaderBuilder()
        .AddWordHandler(new ReaderWordOptions { StreamTextWhenPossible = true }).Build();

    internal static WordReadObservation Read(byte[] fixture) => Read(Reader, fixture);

    /// <summary>The same text through the streaming option, which skips the Word model for simple bodies.</summary>
    internal static WordReadObservation ReadStreamed(byte[] fixture) => Read(StreamingReader, fixture);

    private static WordReadObservation Read(OfficeDocumentReader reader, byte[] fixture) {
        OfficeDocumentReadResult result = reader.ReadDocument(fixture, "fixture.docx", new ReaderOptions {
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
