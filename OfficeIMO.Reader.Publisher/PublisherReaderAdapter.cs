using OfficeIMO.Publisher;

namespace OfficeIMO.Reader.Publisher;

internal static class PublisherReaderAdapter {
    internal static OfficeDocumentReadResult Read(string path, ReaderOptions settings,
        ReaderPublisherOptions options, CancellationToken token) {
        ReaderPublisherOptions operation = options.Clone();
        ReaderAdapterInputSnapshot input = DocumentReaderEngine.ReadAdapterInput(path, settings, token,
            operation.ReadOptions!.Limits.MaxInputBytes);
        return Project(input, settings, operation, token);
    }

    internal static OfficeDocumentReadResult Read(Stream stream, string? name, ReaderOptions settings,
        ReaderPublisherOptions options, CancellationToken token) {
        ReaderPublisherOptions operation = options.Clone();
        ReaderAdapterInputSnapshot input = DocumentReaderEngine.ReadAdapterInput(stream,
            string.IsNullOrWhiteSpace(name) ? "document.pub" : name, settings, token,
            operation.ReadOptions!.Limits.MaxInputBytes);
        return Project(input, settings, operation, token);
    }

    private static OfficeDocumentReadResult Project(ReaderAdapterInputSnapshot input, ReaderOptions settings,
        ReaderPublisherOptions options, CancellationToken token) {
        PublisherDocument document = PublisherDocument.Load(input.Bytes, options.ReadOptions, token);
        OfficeDocumentReadResult result = new PublisherReadProjection(document, input.Source.Path!, settings, options, token).Build();
        result.Source = input.Source;
        foreach (ReaderChunk chunk in result.Chunks) DocumentReaderEngine.ApplyAdapterSource(chunk, input, settings.ComputeHashes);
        return result;
    }
}
