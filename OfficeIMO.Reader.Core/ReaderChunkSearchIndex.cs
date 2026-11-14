using OfficeIMO.Reader.Internal.TextAnalysis;

namespace OfficeIMO.Reader;

/// <summary>
/// In-memory BM25 index over <see cref="ReaderChunk"/> values: chunk text as the body and the heading path as a
/// weighted heading field. Queries accept words, <c>"quoted phrases"</c>, <c>prefix*</c>, <c>AND</c>, <c>OR</c> and
/// parentheses.
/// </summary>
/// <remarks>
/// Words are found with Unicode (UAX #29) word boundaries, normalized to NFKC and case folded. Equal scores rank
/// in the order the chunks were given. Build the index once and search it as often as needed; it keeps references to
/// the chunks, not copies.
/// </remarks>
public sealed class ReaderChunkSearchIndex {
    private readonly ReaderChunk[] _chunks;
    private readonly OfficeTextIndex _index;
    private readonly OfficeDocumentRankingOptions _ranking;

    /// <summary>Indexes <paramref name="chunks"/> in order.</summary>
    public ReaderChunkSearchIndex(IEnumerable<ReaderChunk> chunks, OfficeDocumentRankingOptions? ranking = null) {
        if (chunks == null) throw new ArgumentNullException(nameof(chunks));
        _ranking = ranking ?? new OfficeDocumentRankingOptions();
        _index = ReaderRankedSearch.CreateIndex(_ranking, matchCase: false);
        var indexed = new List<ReaderChunk>();
        foreach (ReaderChunk chunk in chunks) {
            if (chunk == null) throw new ArgumentException("Chunks must not be null.", nameof(chunks));
            if (!ReaderRankedSearch.Add(_index, indexed.Count, chunk.Location?.HeadingPath, chunk.Text)) break;
            indexed.Add(chunk);
        }
        _chunks = indexed.ToArray();
    }

    /// <summary>Number of indexed chunks.</summary>
    public int Count => _chunks.Length;

    /// <summary>True when <see cref="OfficeDocumentRankingOptions.MaximumIndexedTokens"/> cut indexing short.</summary>
    public bool LimitReached => _index.LimitReached;

    /// <summary>Returns up to <paramref name="maximumResults"/> chunks, best first.</summary>
    public ReaderChunkSearchResult Search(string query, int maximumResults = 100) {
        if (query == null) throw new ArgumentNullException(nameof(query));
        if (maximumResults < 1) throw new ArgumentOutOfRangeException(nameof(maximumResults));
        OfficeTextIndexSearchResult result = ReaderRankedSearch.Search(_index, query, _ranking, maximumResults);
        ReaderChunkSearchHit[] hits = result.Hits
            .Select(hit => new ReaderChunkSearchHit(_chunks[hit.DocumentIndex], hit.Score, ReaderRankedSearch.BodyHighlights(hit)))
            .ToArray();
        return new ReaderChunkSearchResult(hits, result.TotalMatches, result.HasSearchableText,
            result.WorkLimitReached || result.PrefixExpansionLimited);
    }
}

/// <summary>Ranked chunks for one query.</summary>
public sealed class ReaderChunkSearchResult {
    internal ReaderChunkSearchResult(IReadOnlyList<ReaderChunkSearchHit> hits, int totalMatches, bool hasSearchableText, bool limitReached) {
        Hits = hits;
        TotalMatches = totalMatches;
        HasSearchableText = hasSearchableText;
        LimitReached = limitReached;
    }

    /// <summary>Hits by descending score, then by chunk order.</summary>
    public IReadOnlyList<ReaderChunkSearchHit> Hits { get; }

    /// <summary>Number of matching chunks before the result limit applied.</summary>
    public int TotalMatches { get; }

    /// <summary>False when the query had no searchable words, for example punctuation only.</summary>
    public bool HasSearchableText { get; }

    /// <summary>True when a prefix expansion or work limit cut evaluation short.</summary>
    public bool LimitReached { get; }
}

/// <summary>One ranked chunk.</summary>
public sealed class ReaderChunkSearchHit {
    internal ReaderChunkSearchHit(ReaderChunk chunk, double score, IReadOnlyList<OfficeDocumentSearchHighlight> highlights) {
        Chunk = chunk;
        Score = score;
        Highlights = highlights;
    }

    /// <summary>The matching chunk.</summary>
    public ReaderChunk Chunk { get; }

    /// <summary>BM25 relevance score; higher is more relevant.</summary>
    public double Score { get; }

    /// <summary>Matched ranges in <see cref="ReaderChunk.Text"/>; empty when only the heading path matched.</summary>
    public IReadOnlyList<OfficeDocumentSearchHighlight> Highlights { get; }
}
