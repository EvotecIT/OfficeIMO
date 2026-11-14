#if OFFICEIMO_READER_CORE
namespace OfficeIMO.Reader.Internal.TextAnalysis;
#else
namespace OfficeIMO.TextAnalysis;
#endif

/// <summary>Ranking, analysis and resource limits for an <see cref="OfficeTextIndex"/>.</summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
sealed class OfficeTextIndexOptions {
    /// <summary>BM25 term-frequency saturation. Must be zero or more. Default: 1.2.</summary>
    public double K1 { get; set; } = 1.2;

    /// <summary>BM25 length normalization, from 0 (none) to 1 (full). Default: 0.75.</summary>
    public double B { get; set; } = 0.75;

    /// <summary>
    /// Field weights by field name (ordinal). Fields not listed weigh 1. A typical setup weighs a title or heading
    /// field above the body.
    /// </summary>
    public IDictionary<string, double> FieldWeights { get; } = new Dictionary<string, double>(StringComparer.Ordinal);

    /// <summary>Analyser settings used for indexed text and query text.</summary>
    public OfficeTextAnalyzerOptions Analyzer { get; set; } = new();

    /// <summary>Maximum number of documents. Further documents are rejected. Default: 1,000,000.</summary>
    public int MaximumDocuments { get; set; } = 1_000_000;

    /// <summary>
    /// Maximum number of tokens held across the index, which bounds postings memory. The document that reaches the
    /// limit is indexed up to it and further documents are rejected. Default: 20,000,000.
    /// </summary>
    public long MaximumTokens { get; set; } = 20_000_000;

    /// <summary>Characters analysed per field; the rest of a longer field is not indexed. Default: 4,000,000.</summary>
    public int MaximumFieldCharacters { get; set; } = 4_000_000;
}

/// <summary>Per-query limits and output settings for <see cref="OfficeTextIndex"/> searches.</summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
sealed class OfficeTextIndexSearchOptions {
    /// <summary>Operator between adjacent clauses when query text is parsed. Default: <see cref="OfficeTextQueryOperator.And"/>.</summary>
    public OfficeTextQueryOperator DefaultOperator { get; set; } = OfficeTextQueryOperator.And;

    /// <summary>Maximum number of ranked hits returned. Default: 100.</summary>
    public int MaximumResults { get; set; } = 100;

    /// <summary>Maximum term, phrase and prefix clauses in one query. Default: 64.</summary>
    public int MaximumClauses { get; set; } = 64;

    /// <summary>
    /// Maximum index terms one prefix expands to. Terms are taken in ordinal order; the result reports when a prefix
    /// had more. Default: 128.
    /// </summary>
    public int MaximumPrefixExpansions { get; set; } = 128;

    /// <summary>
    /// Maximum postings read while evaluating the query. When reached, evaluation stops and the result reports it;
    /// hits found so far are still ranked. Default: 10,000,000.
    /// </summary>
    public long MaximumPostingsVisited { get; set; } = 10_000_000;

    /// <summary>When true (the default), each hit lists the matched source ranges per field.</summary>
    public bool IncludeHighlights { get; set; } = true;
}

/// <summary>One indexed field of a document.</summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
readonly struct OfficeTextField {
    /// <summary>Creates a field.</summary>
    public OfficeTextField(string name, string? text) {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Text = text ?? string.Empty;
    }

    /// <summary>Field name, matched ordinally against <see cref="OfficeTextIndexOptions.FieldWeights"/>.</summary>
    public string Name { get; }

    /// <summary>Field text.</summary>
    public string Text { get; }
}

/// <summary>One matched source range in a field of a hit.</summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
readonly struct OfficeTextHighlight {
    /// <summary>Creates a highlight.</summary>
    public OfficeTextHighlight(string field, int start, int length) {
        Field = field ?? throw new ArgumentNullException(nameof(field));
        Start = start;
        Length = length;
    }

    /// <summary>Field name.</summary>
    public string Field { get; }

    /// <summary>Zero-based UTF-16 offset in the field text.</summary>
    public int Start { get; }

    /// <summary>Length in UTF-16 code units. A phrase covers its first to last word.</summary>
    public int Length { get; }
}

/// <summary>One ranked document.</summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
sealed class OfficeTextIndexHit {
    internal OfficeTextIndexHit(string documentId, int documentIndex, double score, IReadOnlyList<OfficeTextHighlight> highlights) {
        DocumentId = documentId;
        DocumentIndex = documentIndex;
        Score = score;
        Highlights = highlights;
    }

    /// <summary>Identifier given when the document was added.</summary>
    public string DocumentId { get; }

    /// <summary>Zero-based order in which the document was added; ties in score rank by this order.</summary>
    public int DocumentIndex { get; }

    /// <summary>BM25F relevance score; higher is more relevant.</summary>
    public double Score { get; }

    /// <summary>Matched ranges ordered by field (in the document's field order) and offset.</summary>
    public IReadOnlyList<OfficeTextHighlight> Highlights { get; }
}

/// <summary>Ranked results of one query.</summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
sealed class OfficeTextIndexSearchResult {
    internal OfficeTextIndexSearchResult(
        OfficeTextQuery? query,
        IReadOnlyList<OfficeTextIndexHit> hits,
        int totalMatches,
        bool prefixExpansionLimited,
        bool workLimitReached,
        bool hasSearchableText) {
        Query = query;
        HasSearchableText = hasSearchableText;
        Hits = hits;
        TotalMatches = totalMatches;
        PrefixExpansionLimited = prefixExpansionLimited;
        WorkLimitReached = workLimitReached;
    }

    /// <summary>The evaluated query; null when the query text contained no clauses.</summary>
    public OfficeTextQuery? Query { get; }

    /// <summary>
    /// False when no clause had analysable text, for example a query of punctuation only. Callers can then fall back
    /// to literal matching.
    /// </summary>
    public bool HasSearchableText { get; }

    /// <summary>Hits by descending score, then by <see cref="OfficeTextIndexHit.DocumentIndex"/>.</summary>
    public IReadOnlyList<OfficeTextIndexHit> Hits { get; }

    /// <summary>Number of matching documents before <see cref="OfficeTextIndexSearchOptions.MaximumResults"/> applied.</summary>
    public int TotalMatches { get; }

    /// <summary>True when a prefix matched more terms than <see cref="OfficeTextIndexSearchOptions.MaximumPrefixExpansions"/>.</summary>
    public bool PrefixExpansionLimited { get; }

    /// <summary>True when <see cref="OfficeTextIndexSearchOptions.MaximumPostingsVisited"/> stopped evaluation early.</summary>
    public bool WorkLimitReached { get; }
}

/// <summary>A window of field text around the matches of a hit.</summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
sealed class OfficeTextSnippet {
    internal OfficeTextSnippet(string text, int sourceStart, bool startsMidText, bool endsMidText,
        IReadOnlyList<OfficeTextHighlight> highlights) {
        Text = text;
        SourceStart = sourceStart;
        StartsMidText = startsMidText;
        EndsMidText = endsMidText;
        Highlights = highlights;
    }

    /// <summary>Snippet text, cut at grapheme cluster boundaries.</summary>
    public string Text { get; }

    /// <summary>Offset of the snippet in the field text.</summary>
    public int SourceStart { get; }

    /// <summary>True when field text precedes the snippet.</summary>
    public bool StartsMidText { get; }

    /// <summary>True when field text follows the snippet.</summary>
    public bool EndsMidText { get; }

    /// <summary>Highlights inside the snippet, with offsets relative to <see cref="Text"/>.</summary>
    public IReadOnlyList<OfficeTextHighlight> Highlights { get; }
}
