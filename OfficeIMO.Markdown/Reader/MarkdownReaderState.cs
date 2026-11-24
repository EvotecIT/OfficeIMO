namespace OfficeIMO.Markdown;

/// <summary>
/// Reference-style link definition collected while parsing, such as <c>[label]: https://example.com "Title"</c>.
/// </summary>
public sealed class MarkdownReferenceLinkDefinition {
    /// <summary>Normalized reference label used for matching reference-style links and images.</summary>
    public string Label { get; }

    /// <summary>Resolved destination URL after reader profile URL handling.</summary>
    public string Url { get; }

    /// <summary>Optional definition title.</summary>
    public string? Title { get; }

    /// <summary>Source span for the entire reference-style definition, when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Source span for the definition label token, when available.</summary>
    public MarkdownSourceSpan? LabelSourceSpan { get; }

    /// <summary>Source span for the opening <c>[</c> marker before the definition label, when available.</summary>
    public MarkdownSourceSpan? OpeningMarkerSourceSpan { get; }

    /// <summary>Source span for the <c>]:</c> marker after the definition label, when available.</summary>
    public MarkdownSourceSpan? SeparatorMarkerSourceSpan { get; }

    /// <summary>Source span for the destination token, when available.</summary>
    public MarkdownSourceSpan? UrlSourceSpan { get; }

    /// <summary>Source span for the optional title token, when available.</summary>
    public MarkdownSourceSpan? TitleSourceSpan { get; }

    /// <summary>
    /// Creates a reference-style link definition descriptor.
    /// </summary>
    public MarkdownReferenceLinkDefinition(
        string label,
        string url,
        string? title,
        MarkdownSourceSpan? sourceSpan = null,
        MarkdownSourceSpan? labelSourceSpan = null,
        MarkdownSourceSpan? urlSourceSpan = null,
        MarkdownSourceSpan? titleSourceSpan = null,
        MarkdownSourceSpan? openingMarkerSourceSpan = null,
        MarkdownSourceSpan? separatorMarkerSourceSpan = null) {
        Label = label ?? string.Empty;
        Url = url ?? string.Empty;
        Title = title;
        SourceSpan = sourceSpan;
        LabelSourceSpan = labelSourceSpan;
        OpeningMarkerSourceSpan = openingMarkerSourceSpan;
        SeparatorMarkerSourceSpan = separatorMarkerSourceSpan;
        UrlSourceSpan = urlSourceSpan;
        TitleSourceSpan = titleSourceSpan;
    }
}

/// <summary>
/// Markdig-style abbreviation definition collected while parsing, such as
/// <c>*[HTML]: Hyper Text Markup Language</c>.
/// </summary>
public sealed class MarkdownAbbreviationDefinition {
    /// <summary>Source spelling that should be expanded when it appears as a standalone inline token.</summary>
    public string Label { get; }

    /// <summary>Title rendered on the resulting <c>&lt;abbr&gt;</c> element.</summary>
    public string Title { get; }

    /// <summary>Source span for the entire abbreviation definition, when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Source span for the abbreviation label token, when available.</summary>
    public MarkdownSourceSpan? LabelSourceSpan { get; }

    /// <summary>Source span for the opening <c>*[</c> marker, when available.</summary>
    public MarkdownSourceSpan? OpeningMarkerSourceSpan { get; }

    /// <summary>Source span for the closing <c>]:</c> marker, when available.</summary>
    public MarkdownSourceSpan? SeparatorMarkerSourceSpan { get; }

    /// <summary>Source span for the title token, when available.</summary>
    public MarkdownSourceSpan? TitleSourceSpan { get; }

    /// <summary>Whether the definition was authored as leading content inside a list item.</summary>
    public bool IsListItemDefinition { get; }

    // Zero-based source line and its text, kept without a syntax tree so a semantic parse orders
    // and writes parse-owned definitions as the source-backed parse does.
    internal int SourceLineIndex { get; set; } = -1;
    internal string? SourceLineText { get; set; }

    /// <summary>Creates an abbreviation definition descriptor.</summary>
    public MarkdownAbbreviationDefinition(
        string label,
        string title,
        MarkdownSourceSpan? sourceSpan = null,
        MarkdownSourceSpan? labelSourceSpan = null,
        MarkdownSourceSpan? titleSourceSpan = null,
        MarkdownSourceSpan? openingMarkerSourceSpan = null,
        MarkdownSourceSpan? separatorMarkerSourceSpan = null,
        bool isListItemDefinition = false) {
        Label = label ?? string.Empty;
        Title = title ?? string.Empty;
        SourceSpan = sourceSpan;
        LabelSourceSpan = labelSourceSpan;
        OpeningMarkerSourceSpan = openingMarkerSourceSpan;
        SeparatorMarkerSourceSpan = separatorMarkerSourceSpan;
        TitleSourceSpan = titleSourceSpan;
        IsListItemDefinition = isListItemDefinition;
    }
}

/// <summary>
/// Mutable per-parse state shared across block and inline parsers.
/// </summary>
public sealed class MarkdownReaderState {
    // Most documents define neither, so the dictionaries are created by the first definition.
    private static readonly Dictionary<string, MarkdownReferenceLinkDefinition> EmptyLinkRefs = new(System.StringComparer.OrdinalIgnoreCase);
    private static readonly Dictionary<string, MarkdownAbbreviationDefinition> EmptyAbbreviations = new(System.StringComparer.Ordinal);
    private Dictionary<string, MarkdownReferenceLinkDefinition>? _linkRefs;
    private Dictionary<string, MarkdownAbbreviationDefinition>? _abbreviations;
    // Nested parses start from their parent's definitions. Copying them for every nested container
    // made documents with many definitions and many containers quadratic, so states share the
    // dictionaries and whichever side changes them first takes a copy of its own.
    private bool _linkRefsShared;
    private bool _abbreviationsShared;

    /// <summary>Creates an empty parse state.</summary>
    public MarkdownReaderState() {
    }

    private MarkdownReaderState(
        Dictionary<string, MarkdownReferenceLinkDefinition>? linkRefs,
        Dictionary<string, MarkdownAbbreviationDefinition>? abbreviations) {
        _linkRefs = linkRefs;
        _abbreviations = abbreviations;
    }

    /// <summary>Creates a state that starts from the definitions of <paramref name="source"/> without copying them.</summary>
    internal static MarkdownReaderState CreateSharingDefinitionsWith(MarkdownReaderState source) {
        var state = new MarkdownReaderState(source._linkRefs, source._abbreviations);
        state._linkRefsShared = source._linkRefsShared = true;
        state._abbreviationsShared = source._abbreviationsShared = true;
        return state;
    }

    /// <summary>Reference-style link definitions collected while parsing. Change them through <see cref="EditLinkRefs"/>.</summary>
    internal IReadOnlyDictionary<string, MarkdownReferenceLinkDefinition> LinkRefs => _linkRefs ?? EmptyLinkRefs;
    /// <summary>Case-sensitive abbreviation definitions collected while parsing. Later definitions replace earlier ones, matching Markdig. Change them through <see cref="EditAbbreviations"/>.</summary>
    internal IReadOnlyDictionary<string, MarkdownAbbreviationDefinition> Abbreviations => _abbreviations ?? EmptyAbbreviations;

    internal Dictionary<string, MarkdownReferenceLinkDefinition> EditLinkRefs() {
        if (_linkRefs == null) {
            _linkRefs = new Dictionary<string, MarkdownReferenceLinkDefinition>(System.StringComparer.OrdinalIgnoreCase);
            _linkRefsShared = false;
        } else if (_linkRefsShared) {
            _linkRefs = CopyDefinitions(_linkRefs, System.StringComparer.OrdinalIgnoreCase);
            _linkRefsShared = false;
        }

        return _linkRefs;
    }

    internal Dictionary<string, MarkdownAbbreviationDefinition> EditAbbreviations() {
        if (_abbreviations == null) {
            _abbreviations = new Dictionary<string, MarkdownAbbreviationDefinition>(System.StringComparer.Ordinal);
            _abbreviationsShared = false;
        } else if (_abbreviationsShared) {
            _abbreviations = CopyDefinitions(_abbreviations, System.StringComparer.Ordinal);
            _abbreviationsShared = false;
        }

        return _abbreviations;
    }

    private MarkdownReaderParseCache? _parseCache;

    /// <summary>Values derived from the parse options, shared by this state and the states cloned from it.</summary>
    internal MarkdownReaderParseCache ParseCache => _parseCache ??= new MarkdownReaderParseCache();

    internal void ShareParseCacheWith(MarkdownReaderState source) {
        _parseCache = source.ParseCache;
    }


    private static Dictionary<string, TValue> CopyDefinitions<TValue>(Dictionary<string, TValue> source, System.StringComparer comparer) {
        // Enumeration order is kept, as copying entry by entry always did.
        var copy = new Dictionary<string, TValue>(source.Count, comparer);
        foreach (var pair in source) {
            copy[pair.Key] = pair.Value;
        }

        return copy;
    }
    internal int SourceLineOffset { get; set; }
    // The quoted content of the run of quote lines last measured for HTML blocks, shared by sibling quotes and callouts.
    internal QuotedContentRun? QuotedContentRun { get; set; }
    // A source-line slice is already a real line, even when its text is empty.
    internal bool LastInputLineIsContent { get; set; }
    internal bool HasSyntheticTerminalLine { get; set; }
    internal MarkdownSourceTextMap? SourceTextMap { get; set; }
    internal int ListMarkerIndentOffset { get; set; }
    internal bool UseRelativeListMarkerIndent { get; set; }
    /// <summary>Whether this parse is producing the source-backed syntax representation.</summary>
    internal bool CaptureSyntaxTree { get; set; }
    /// <summary>Whether a projection parse should retain only top-level block source spans.</summary>
    internal bool CaptureBlockSourceSpans { get; set; }
    /// <summary>Whether table parsing should build the rich inline and structured cell model.</summary>
    internal bool BuildTableCellModels { get; set; } = true;
    /// <summary>Enables a lightweight semantic pass used only to locate an exact nested-list boundary.</summary>
    internal bool IsListBoundaryProbe { get; set; }
    // Set when list continuation collection accepts a lazy continuation line.
    internal bool LazyListContinuationSeen { get; set; }
    internal bool SuppressBlockGenericAttributes { get; set; }
    internal bool SuppressHeadingGenericAttributes { get; set; }
    internal bool IsMarkdigDefinitionListBody { get; set; }
    internal int NestedBlockDepth { get; set; }
    private List<MarkdownDocument>? _nestedListScratchDocuments;
    private int _nestedListScratchDepth;

    /// <summary>Takes the holder document for a nested list parsed without syntax capture.</summary>
    internal MarkdownDocument RentNestedListScratchDocument() {
        var documents = _nestedListScratchDocuments ??= new List<MarkdownDocument>();
        if (_nestedListScratchDepth == documents.Count) {
            documents.Add(MarkdownDocument.Create());
        }

        return documents[_nestedListScratchDepth++];
    }

    internal void ReturnNestedListScratchDocument(MarkdownDocument document) {
        document.ClearParserScratchBlocks();
        _nestedListScratchDepth--;
    }
    internal bool DiscardsNestedListSourceRemap { get; set; }
    internal MarkdownPendingGenericAttributeBlock? PendingGenericAttributeBlock { get; set; }
    internal MarkdownLineIndexSet LazyQuoteContinuationLines;
    internal MarkdownLineIndexSet QuoteContainerLines;
    internal MarkdownLineIndexSet SuppressedSetextHeadingUnderlineLines;
    internal MarkdownLineIndexSet SuppressedParagraphGenericAttributeStartLines;
    internal IReadOnlyList<int>? SourceLineAbsoluteNumbers { get; set; }
}

/// <summary>
/// A set of zero-based line indexes that allocates its storage on the first added line. It is a mutable
/// struct held in fields of <see cref="MarkdownReaderState"/>: change it through the field, never a copy.
/// </summary>
internal struct MarkdownLineIndexSet {
    private HashSet<int>? _lines;

    internal int Count => _lines?.Count ?? 0;

    internal bool Contains(int line) => _lines != null && _lines.Contains(line);

    internal void Add(int line) => (_lines ??= new HashSet<int>()).Add(line);

    internal void Clear() => _lines?.Clear();

    internal void CopyTo(ref MarkdownLineIndexSet target) {
        if (_lines == null) {
            return;
        }

        foreach (int line in _lines) {
            target.Add(line);
        }
    }
}

internal sealed class MarkdownPendingGenericAttributeBlock {
    internal MarkdownPendingGenericAttributeBlock(
        MarkdownAttributeSet attributes,
        string sourceText,
        MarkdownSourceSpan? sourceSpan) {
        Attributes = attributes ?? MarkdownAttributeSet.Empty;
        SourceText = sourceText ?? string.Empty;
        SourceSpan = sourceSpan;
    }

    internal MarkdownAttributeSet Attributes { get; }
    internal string SourceText { get; }
    internal MarkdownSourceSpan? SourceSpan { get; }
}
