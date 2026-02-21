namespace OfficeIMO.Reader.Html;

/// <summary>
/// Registration helpers for plugging HTML support into <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderHtmlRegistrationExtensions {
    /// <summary>
    /// Stable handler identifier for HTML adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.html";

    /// <summary>
    /// Registers HTML ingestion into <see cref="DocumentReader"/> for <c>.html</c> and <c>.htm</c>.
    /// </summary>
    public static void RegisterHtmlHandler(bool replaceExisting = false) {
        DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "HTML Reader Adapter",
            Description = "Modular HTML adapter using OfficeIMO.Word.Html + OfficeIMO.Word.Markdown.",
            Kind = ReaderInputKind.Unknown,
            Extensions = new[] { ".html", ".htm" },
            ReadPath = (path, readerOptions, ct) => DocumentReaderHtmlExtensions.ReadHtmlFile(
                htmlPath: path,
                readerOptions: readerOptions,
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => ReadHtmlStream(stream, sourceName, readerOptions, ct)
        }, replaceExisting);
    }

    /// <summary>
    /// Unregisters HTML ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterHtmlHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }

    private static IEnumerable<ReaderChunk> ReadHtmlStream(Stream stream, string? sourceName, ReaderOptions options, CancellationToken ct) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        var sb = new StringBuilder();
        var buffer = new char[16 * 1024];
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 16 * 1024, leaveOpen: true);

        while (true) {
            ct.ThrowIfCancellationRequested();
            var read = reader.Read(buffer, 0, buffer.Length);
            if (read <= 0) break;
            sb.Append(buffer, 0, read);
        }

        var name = string.IsNullOrWhiteSpace(sourceName) ? "document.html" : sourceName!;
        return DocumentReaderHtmlExtensions.ReadHtmlString(
            html: sb.ToString(),
            sourceName: name,
            readerOptions: options,
            cancellationToken: ct);
    }
}
