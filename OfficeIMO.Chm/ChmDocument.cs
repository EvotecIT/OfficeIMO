using OfficeIMO.Core.Internal;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Chm;

/// <summary>An immutable, cross-platform snapshot of a compiled HTML help book. Loading executes no help content.</summary>
public sealed partial class ChmDocument {
    private readonly ChmReadOptions _options;
    private readonly Dictionary<string, ChmEntry> _entries;
    private readonly List<ChmDiagnostic> _diagnostics = new List<ChmDiagnostic>();
    private Encoding _encoding = Encoding.UTF8;
    private string? _contentsPath;
    private string? _indexPath;
    internal ChmDocument(IReadOnlyList<ChmEntry> entries, ChmReadOptions options, uint version, uint locale, CancellationToken token) {
        _options = options; Version = version; LocaleId = locale; Entries = entries;
        _entries = entries.ToDictionary(entry => entry.Path, StringComparer.OrdinalIgnoreCase);
        Resources = Array.AsReadOnly(entries.Where(entry => !entry.IsSystem && !entry.IsDirectory).ToArray());
        ReadMetadata(token);
        var navigation = new ChmNavigationReader(this, _options, _encoding, token);
        TableOfContents = navigation.ReadContents(_contentsPath);
        Index = navigation.ReadIndex(_indexPath);
        var titles = navigation.TopicTitles;
        var ordered = new List<ChmEntry>();
        var included = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (ChmNavigationItem item in EnumerateNavigation(TableOfContents)) {
            foreach (ChmLink link in item.Links) {
                ChmEntry? entry = FindEntry(link.Target);
                if (entry == null || !IsTopic(entry)) continue;
                if (included.Add(entry.Path)) ordered.Add(entry);
                if (!titles.ContainsKey(entry.Path)) titles.Add(entry.Path, link.Title ?? item.Name);
            }
        }
        if (ordered.Count == 0 && DefaultTopic != null) {
            ChmEntry? entry = FindEntry(DefaultTopic);
            if (entry != null && IsTopic(entry) && included.Add(entry.Path)) ordered.Add(entry);
        }
        foreach (ChmEntry entry in Resources.OrderBy(entry => entry.Path, StringComparer.Ordinal)) {
            token.ThrowIfCancellationRequested();
            if (IsTopic(entry) && included.Add(entry.Path)) ordered.Add(entry);
        }
        Topics = Array.AsReadOnly(ordered.Select(entry => new ChmTopic(this, entry,
            titles.TryGetValue(entry.Path, out string? title) && !string.IsNullOrWhiteSpace(title) ? title : entry.Path)).ToArray());
        Diagnostics = _diagnostics.AsReadOnly();
    }

    /// <summary>Loads a file into a bounded independent snapshot.</summary>
    public static ChmDocument Load(string path, ChmReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("CHM path cannot be empty.", nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Load(stream, options, cancellationToken);
    }
    /// <summary>Loads caller-owned input. Seekable input is rewound and its original position restored; forward-only input is read from its current position.</summary>
    public static ChmDocument Load(Stream stream, ChmReadOptions? options = null, CancellationToken cancellationToken = default) {
        ChmReadOptions configured = Configure(options);
        byte[] bytes;
        try { bytes = OfficeStreamReader.ReadAllBytes(stream, cancellationToken, configured.MaxInputBytes); }
        catch (InvalidDataException exception) when (OfficeStreamReader.IsSizeLimitException(exception)) {
            throw new ChmReadException("CHM_INPUT_LIMIT", "The CHM archive exceeds MaxInputBytes.", exception);
        }
        return ChmArchiveReader.Read(bytes, configured, cancellationToken);
    }
    /// <summary>Loads an independent snapshot of the supplied archive bytes.</summary>
    public static ChmDocument Load(byte[] bytes, ChmReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        ChmReadOptions configured = Configure(options);
        cancellationToken.ThrowIfCancellationRequested();
        if (bytes.LongLength > configured.MaxInputBytes) throw ChmBinary.Error("INPUT_LIMIT", "The CHM archive exceeds MaxInputBytes.");
        return ChmArchiveReader.Read((byte[])bytes.Clone(), configured, cancellationToken);
    }
    /// <summary>Asynchronously snapshots a file and parses it with cooperative cancellation.</summary>
    public static async Task<ChmDocument> LoadAsync(string path, ChmReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("CHM path cannot be empty.", nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, true);
        return await LoadAsync(stream, options, cancellationToken).ConfigureAwait(false);
    }
    /// <summary>Asynchronously snapshots caller-owned input; parsing and LZX decoding observe the same cancellation token.</summary>
    public static async Task<ChmDocument> LoadAsync(Stream stream, ChmReadOptions? options = null, CancellationToken cancellationToken = default) {
        ChmReadOptions configured = Configure(options);
        byte[] bytes;
        try { bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, configured.MaxInputBytes).ConfigureAwait(false); }
        catch (InvalidDataException exception) when (OfficeStreamReader.IsSizeLimitException(exception)) {
            throw new ChmReadException("CHM_INPUT_LIMIT", "The CHM archive exceeds MaxInputBytes.", exception);
        }
        return ChmArchiveReader.Read(bytes, configured, cancellationToken);
    }

    /// <summary>ITSF container version.</summary>
    public uint Version { get; }
    /// <summary>Help-book LCID from #SYSTEM, falling back to the ITSF header.</summary>
    public uint LocaleId { get; private set; }
    /// <summary>Title declared in #SYSTEM, or null when absent.</summary>
    public string? Title { get; private set; }
    /// <summary>Default topic reference, including any fragment.</summary>
    public string? DefaultTopic { get; private set; }
    /// <summary>Compiler identification declared in #SYSTEM.</summary>
    public string? Compiler { get; private set; }
    /// <summary>Every archive entry, including internal metadata, in directory order.</summary>
    public IReadOnlyList<ChmEntry> Entries { get; }
    /// <summary>All content files, including topics and sitemaps, excluding internal streams and directory markers.</summary>
    public IReadOnlyList<ChmEntry> Resources { get; }
    /// <summary>HTML topics in contents order followed by unlisted topics in ordinal path order.</summary>
    public IReadOnlyList<ChmTopic> Topics { get; }
    /// <summary>Contents hierarchy from compiled tables or the declared HTML sitemap.</summary>
    public IReadOnlyList<ChmNavigationItem> TableOfContents { get; }
    /// <summary>Keyword index, including multiple targets and See Also references.</summary>
    public IReadOnlyList<ChmNavigationItem> Index { get; }
    /// <summary>Non-fatal metadata and unresolved-reference diagnostics.</summary>
    public IReadOnlyList<ChmDiagnostic> Diagnostics { get; }

    /// <summary>Finds a content reference relative to a topic or sitemap. External, escaping, and merged-help references return null.</summary>
    public ChmEntry? FindEntry(string reference, string sourcePath = "/") {
        if (reference == null) throw new ArgumentNullException(nameof(reference));
        if (sourcePath == null) throw new ArgumentNullException(nameof(sourcePath));
        if (_entries.TryGetValue(reference.Replace('\\', '/'), out ChmEntry? exact)) return exact;
        string? path = reference.StartsWith("::", StringComparison.Ordinal) ? reference : ChmPaths.Resolve(reference, sourcePath, _options.MaxPathLength);
        return path != null && _entries.TryGetValue(path, out ChmEntry? entry) ? entry : null;
    }

    internal string ReadHtml(ChmEntry entry, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        using Stream stream = entry.OpenRead();
        Encoding encoding = new HtmlTextEncodingResolver(_options.EncodingProvider).ResolveHtmlEncoding(stream, _options.TextEncoding, _encoding);
        using var reader = new StreamReader(stream, encoding, _options.TextEncoding == null, 4096, true);
        var text = new StringBuilder();
        var buffer = new char[4096];
        int read;
        while ((read = reader.Read(buffer, 0, buffer.Length)) > 0) { token.ThrowIfCancellationRequested(); text.Append(buffer, 0, read); }
        return text.ToString();
    }
    internal HtmlConversionDocumentOptions CreateHtmlOptions(string path, HtmlConversionDocumentOptions? requested = null) {
        var options = requested?.Clone() ?? new HtmlConversionDocumentOptions {
            ParserProvider = _options.ParserProvider, InputEncodingProvider = _options.EncodingProvider
        };
        options.BaseUri = GetTopicUri(path);
        if (options.NormalizationOptions != null) {
            // Resolve authored bases from the topic URI, never from caller fallback context.
            options.NormalizationOptions.BaseUri = null;
            options.NormalizationOptions.BaseElementBaseUri = null;
        }
        options.Limits.MaxInputCharacters = Math.Min(requested?.Limits.MaxInputCharacters ?? int.MaxValue, _options.MaxEntryBytes);
        options.Limits.MaxHtmlNodes = Math.Min(requested?.Limits.MaxHtmlNodes ?? int.MaxValue, _options.MaxHtmlNodes);
        options.Limits.MaxHtmlDepth = Math.Min(requested?.Limits.MaxHtmlDepth ?? int.MaxValue, _options.MaxHtmlDepth);
        options.UrlPolicy.AllowedUrlSchemes.Add("chm");
        options.ResourceUrlPolicy.AllowedUrlSchemes.Add("chm");
        return options;
    }
    internal HtmlDocument ParseSitemap(ChmEntry entry, CancellationToken token) => _options.ParserProvider.ParseDocument(ReadHtml(entry, token),
        new HtmlParseOptions { MaxInputCharacters = _options.MaxEntryBytes, MaxNodes = _options.MaxHtmlNodes, MaxDepth = _options.MaxHtmlDepth }, token);
    internal void Diagnostic(string code, string message, string? path = null) => _diagnostics.Add(new ChmDiagnostic(code, message, path));
    internal static bool IsTopic(ChmEntry entry) => !entry.IsSystem &&
        (entry.Path.EndsWith(".htm", StringComparison.OrdinalIgnoreCase) || entry.Path.EndsWith(".html", StringComparison.OrdinalIgnoreCase) || entry.Path.EndsWith(".xhtml", StringComparison.OrdinalIgnoreCase));
    private static ChmReadOptions Configure(ChmReadOptions? options) { ChmReadOptions configured = options?.Clone() ?? new ChmReadOptions(); configured.Validate(); return configured; }
    internal static IEnumerable<ChmNavigationItem> EnumerateNavigation(IReadOnlyList<ChmNavigationItem> roots) {
        var stack = new Stack<ChmNavigationItem>(roots.Reverse());
        while (stack.Count != 0) { ChmNavigationItem item = stack.Pop(); yield return item; for (int i = item.Children.Count - 1; i >= 0; i--) stack.Push(item.Children[i]); }
    }
}
