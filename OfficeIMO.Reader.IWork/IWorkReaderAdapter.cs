using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.Reader.IWork;

internal static class IWorkReaderAdapter {
    internal static bool Probe(Stream stream, string? sourceName, ReaderOptions readerOptions,
        ReaderIWorkOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(sourceName) || !stream.CanSeek) return false;
        if (!ExpectedKind(sourceName!).HasValue) return false;
        IWorkReadOptions readOptions = options.ReadOptions ?? new IWorkReadOptions();
        long maximumPackageBytes = Math.Min(readOptions.MaximumPackageBytes,
            readerOptions.MaxInputBytes ?? readOptions.MaximumPackageBytes);
        return IWorkContainerProbe.HasModernIndex(stream, maximumPackageBytes,
            readOptions.MaximumEntryCount, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions readerOptions,
        ReaderIWorkOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        IWorkDocumentKind? expected = ExpectedKind(path);
        IWorkSourceDocument source = expected.HasValue
            ? IWorkSourceDocument.Open(path, expected.Value, options.ReadOptions, cancellationToken)
            : IWorkSourceDocument.Open(path, options.ReadOptions, cancellationToken);
        return Project(source, path, readerOptions, options, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(Stream stream, string? sourceName,
        ReaderOptions readerOptions, ReaderIWorkOptions options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        string logicalName = string.IsNullOrWhiteSpace(sourceName)
            ? "document.pages" : sourceName!.Trim();
        IWorkDocumentKind? expected = ExpectedKind(logicalName);
        long originalPosition = stream.CanSeek ? stream.Position : 0;
        try {
            IWorkSourceDocument source = expected.HasValue
                ? IWorkSourceDocument.Open(stream, expected.Value, options.ReadOptions, cancellationToken)
                : IWorkSourceDocument.Open(stream, options.ReadOptions, cancellationToken);
            return Project(source, logicalName, readerOptions, options, cancellationToken);
        } finally {
            if (stream.CanSeek) stream.Position = originalPosition;
        }
    }

    private static IWorkDocumentKind? ExpectedKind(string name) =>
        Path.GetExtension(name).ToLowerInvariant() switch {
            ".pages" => IWorkDocumentKind.Pages,
            ".numbers" => IWorkDocumentKind.Numbers,
            ".key" => IWorkDocumentKind.Keynote,
            _ => null
        };

    private static OfficeDocumentReadResult Project(IWorkSourceDocument source, string path,
        ReaderOptions readerOptions, ReaderIWorkOptions options,
        CancellationToken cancellationToken) {
        var result = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.IWork,
            Source = new OfficeDocumentSource { Path = path },
            CapabilitiesUsed = new[] { "officeimo.reader.iwork", "officeimo.iwork.semantic-source" }
        };
        var projection = new IWorkReadProjection(result, path, readerOptions, options,
            cancellationToken);
        switch (source.Kind) {
            case IWorkDocumentKind.Pages:
                projection.AddPages(source.ReadPages());
                break;
            case IWorkDocumentKind.Numbers:
                projection.AddNumbers(source.ReadNumbers());
                break;
            case IWorkDocumentKind.Keynote:
                projection.AddKeynote(source.ReadKeynote());
                break;
            default:
                throw new InvalidDataException("Unsupported iWork document kind.");
        }
        cancellationToken.ThrowIfCancellationRequested();
        projection.Complete(source);
        return result;
    }
}
