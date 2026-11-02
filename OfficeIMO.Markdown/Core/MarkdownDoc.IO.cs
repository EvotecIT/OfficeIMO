using OfficeIMO.Core.Internal;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Markdown;

public partial class MarkdownDocument : IDisposable {
    private static readonly Encoding Utf8WithoutBom = new UTF8Encoding(false);
    private string? _filePath;

    /// <summary>Gets the associated path, or null for an in-memory document.</summary>
    public string? FilePath => _filePath;
    /// <summary>Gets the source syntax and transform snapshot, when the document was parsed.</summary>
    public MarkdownParseInfo? ParseInfo => _parseResult;

    /// <summary>Parses Markdown text into the document model.</summary>
    public static MarkdownDocument Parse(string markdown, MarkdownLoadOptions? options = null) =>
        options?.CaptureSyntaxTree == false ? MarkdownReader.ParseSemantic(markdown, options) : MarkdownReader.Parse(markdown, options);

    /// <summary>Parses Markdown text and reports the completed native read.</summary>
    public static MarkdownDocument Parse(string markdown, MarkdownLoadOptions? options,
        out OfficeConversionReport report) {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        MarkdownDocument value = Parse(markdown, options);
        report = new OfficeConversionReport("Markdown", "Markdown", itemsWritten: value.Blocks.Count, elapsed: timer.Elapsed);
        return value;
    }

    /// <summary>Loads Markdown from a path and associates it for subsequent saves.</summary>
    public static MarkdownDocument Load(string path, MarkdownLoadOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        MarkdownDocument document = Load(stream, options);
        document._filePath = Path.GetFullPath(path);
        return document;
    }

    /// <summary>Loads Markdown from a caller-owned stream, preserving seekable stream position.</summary>
    public static MarkdownDocument Load(Stream stream, MarkdownLoadOptions? options = null) =>
        Parse(DecodeText(OfficeStreamReader.ReadAllBytes(stream, options?.MaxInputBytes), options?.Encoding ?? Utf8WithoutBom), options);

    /// <summary>Loads a Markdown file and returns its completed native read report.</summary>
    public static MarkdownDocument Load(string path, MarkdownLoadOptions? options, out OfficeConversionReport report) {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        MarkdownDocument document = Load(path, options);
        report = new OfficeConversionReport(Path.GetFullPath(path), "Markdown", itemsWritten: document.Blocks.Count, elapsed: timer.Elapsed);
        return document;
    }

    /// <summary>Loads a caller-owned Markdown stream and returns its completed native read report.</summary>
    public static MarkdownDocument Load(Stream stream, MarkdownLoadOptions? options, out OfficeConversionReport report) {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        MarkdownDocument document = Load(stream, options);
        report = new OfficeConversionReport("Markdown", "Markdown", itemsWritten: document.Blocks.Count, elapsed: timer.Elapsed);
        return document;
    }

    /// <summary>Asynchronously loads a Markdown file and associates its path.</summary>
    public static async Task<MarkdownDocument> LoadAsync(string path, MarkdownLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, true);
        MarkdownDocument document = await LoadAsync(stream, options, cancellationToken).ConfigureAwait(false);
        document._filePath = Path.GetFullPath(path);
        return document;
    }

    /// <summary>Asynchronously loads Markdown from a caller-owned stream.</summary>
    public static async Task<MarkdownDocument> LoadAsync(Stream stream, MarkdownLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = (options ?? new MarkdownLoadOptions()).Clone();
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, options?.MaxInputBytes).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return Parse(DecodeText(bytes, options?.Encoding ?? Utf8WithoutBom), options);
    }

    /// <summary>Encodes the native document with the selected writer and encoding.</summary>
    public byte[] ToBytes(MarkdownSaveOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureNativeSaveAllowed(options);
        return (options?.Encoding ?? Utf8WithoutBom).GetBytes(ToMarkdown(options, cancellationToken));
    }

    /// <summary>Saves to the associated path.</summary>
    public void Save() => Save(_filePath ?? throw new InvalidOperationException("The document has no associated path."));
    /// <summary>Saves to the associated path with explicit writer options.</summary>
    public void Save(MarkdownSaveOptions options) => Save(_filePath ?? throw new InvalidOperationException("The document has no associated path."), options);

    /// <summary>Atomically saves the native document and associates its path.</summary>
    public void Save(string path, MarkdownSaveOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        OfficeFileCommit.WriteAllBytes(path, ToBytes(options));
        _filePath = Path.GetFullPath(path);
    }

    /// <summary>Writes native Markdown to a caller-owned stream.</summary>
    public void Save(Stream stream, MarkdownSaveOptions? options = null) =>
        OfficeStreamWriter.WriteAllBytes(stream, ToBytes(options));

    /// <summary>Asynchronously saves to the associated path.</summary>
    public Task SaveAsync(CancellationToken cancellationToken = default) =>
        SaveAsync(_filePath ?? throw new InvalidOperationException("The document has no associated path."), null, cancellationToken);

    /// <summary>Asynchronously saves the native document and associates its path.</summary>
    public async Task SaveAsync(string path, MarkdownSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        await OfficeFileCommit.WriteAllBytesAsync(path, ToBytes(options, cancellationToken), cancellationToken: cancellationToken).ConfigureAwait(false);
        _filePath = Path.GetFullPath(path);
    }

    /// <summary>Asynchronously writes native Markdown to a caller-owned stream.</summary>
    public Task SaveAsync(Stream stream, MarkdownSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeStreamWriter.WriteAllBytesAsync(stream, ToBytes(options, cancellationToken), cancellationToken);

    private void EnsureNativeSaveAllowed(MarkdownSaveOptions? options) {
        if (options?.AllowLoss != true) AssessNativeSave(options).ThrowIfLoss();
    }

    internal OfficeConversionReport AssessNativeSave(MarkdownSaveOptions? options) {
        var diagnostics = new List<OfficeDiagnostic>();
        if (options?.FrontMatterRendering == MarkdownFrontMatterRenderingMode.Omit && HasDocumentHeader)
            diagnostics.Add(new OfficeDiagnostic(OfficeDiagnosticCodes.ContentOmitted, OfficeDiagnosticSeverity.Warning,
                "Native save would omit document front matter.", "FrontMatter", "Markdown", OfficeConversionLossKind.Omission));
        if (options?.AbbreviationDefinitionRendering == MarkdownAbbreviationDefinitionRenderingMode.Omit && _parseResult?.AbbreviationDefinitions.Count > 0)
            diagnostics.Add(new OfficeDiagnostic(OfficeDiagnosticCodes.ContentOmitted, OfficeDiagnosticSeverity.Warning,
                "Native save would omit abbreviation definitions.", "Abbreviations", "Markdown", OfficeConversionLossKind.Omission));
        return new OfficeConversionReport("Markdown", "Markdown", diagnostics: diagnostics);
    }

    private static string DecodeText(byte[] bytes, Encoding encoding) {
        using var stream = new MemoryStream(bytes, writable: false);
        using var reader = new StreamReader(stream, encoding, true, 1024, false);
        return reader.ReadToEnd();
    }

    /// <summary>Completes this in-memory document's lifetime without closing caller-owned streams.</summary>
    public void Dispose() { }
    /// <summary>Completes this in-memory document's lifetime asynchronously.</summary>
    public Task DisposeAsync() { Dispose(); return Task.CompletedTask; }
}
