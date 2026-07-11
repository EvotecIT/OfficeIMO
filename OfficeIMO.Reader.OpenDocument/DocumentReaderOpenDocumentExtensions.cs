using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

/// <summary>Native OpenDocument ingestion adapter for <see cref="DocumentReader"/>.</summary>
public static partial class DocumentReaderOpenDocumentExtensions {
    /// <summary>Reads an ODT, ODS, or ODP file and emits format-aligned chunks.</summary>
    public static IEnumerable<ReaderChunk> ReadOpenDocument(string path, ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (!File.Exists(path)) throw new FileNotFoundException($"OpenDocument file '{path}' doesn't exist.", path);
        ReaderOptions effective = options ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, effective.MaxInputBytes);
        using OdfDocument document = OdfDocument.OpenAny(path, CreateOpenOptions(effective));
        foreach (ReaderChunk chunk in ReadDocument(document, Path.GetFullPath(path), effective, cancellationToken)) yield return chunk;
    }

    /// <summary>Reads an ODT, ODS, or ODP stream without closing the caller's stream.</summary>
    public static IEnumerable<ReaderChunk> ReadOpenDocument(Stream stream, string? sourceName = null, ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("OpenDocument stream must be readable.", nameof(stream));
        ReaderOptions effective = options ?? new ReaderOptions();
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(stream, effective.MaxInputBytes, cancellationToken, out bool ownsStream);
        try {
            using OdfDocument document = OdfDocument.OpenAny(parseStream, CreateOpenOptions(effective));
            string logicalName = string.IsNullOrWhiteSpace(sourceName) ? "document.odf" : sourceName!.Trim();
            foreach (ReaderChunk chunk in ReadDocument(document, logicalName, effective, cancellationToken)) yield return chunk;
        } finally {
            if (ownsStream) parseStream.Dispose();
        }
    }

    private static IEnumerable<ReaderChunk> ReadDocument(OdfDocument document, string sourceName, ReaderOptions options,
        CancellationToken cancellationToken) {
        if (document is OdtDocument text) {
            foreach (ReaderChunk chunk in ReadTextDocument(text, sourceName, options, cancellationToken)) yield return chunk;
        } else if (document is OdsDocument spreadsheet) {
            foreach (ReaderChunk chunk in ReadSpreadsheet(spreadsheet, sourceName, options, cancellationToken)) yield return chunk;
        } else if (document is OdpPresentation presentation) {
            foreach (ReaderChunk chunk in ReadPresentation(presentation, sourceName, options, cancellationToken)) yield return chunk;
        }
    }

    private static OdfOpenOptions CreateOpenOptions(ReaderOptions options) {
        var result = new OdfOpenOptions();
        if (options.MaxInputBytes.HasValue) result.MaxPackageBytes = options.MaxInputBytes.Value;
        if (options.OpenXmlMaxCharactersInPart.HasValue) result.MaxXmlCharacters = options.OpenXmlMaxCharactersInPart.Value;
        return result;
    }

    private static IEnumerable<string> SplitText(string text, int maxChars) {
        if (string.IsNullOrWhiteSpace(text)) yield break;
        int size = maxChars > 0 ? maxChars : 8000;
        int index = 0;
        while (index < text.Length) {
            int length = Math.Min(size, text.Length - index);
            int end = index + length;
            if (end < text.Length) {
                int split = text.LastIndexOf(' ', end - 1, length);
                if (split > index + Math.Min(128, size / 4)) end = split;
            }
            string piece = text.Substring(index, end - index).Trim();
            if (piece.Length > 0) yield return piece;
            index = end;
            while (index < text.Length && char.IsWhiteSpace(text[index])) index++;
        }
    }

    private static string BuildId(string sourceName, string kind, int index, int part = 0) {
        string safe = new string(Path.GetFileNameWithoutExtension(sourceName).ToLowerInvariant()
            .Select(character => char.IsLetterOrDigit(character) ? character : '-')
            .ToArray()).Trim('-');
        if (safe.Length == 0) safe = "document";
        return string.Concat("odf-", safe, "-", kind, "-", index.ToString("D4", CultureInfo.InvariantCulture), "-", part.ToString("D3", CultureInfo.InvariantCulture));
    }
}
