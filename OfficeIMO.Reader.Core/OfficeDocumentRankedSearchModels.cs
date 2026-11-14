namespace OfficeIMO.Reader;

/// <summary>How <see cref="OfficeDocumentReadResultExtensions.Search"/> matches the query.</summary>
public enum OfficeDocumentSearchMode {
    /// <summary>Every literal occurrence of the query text, in source order.</summary>
    Literal,
    /// <summary>
    /// BM25-ranked word search with UAX #29 word boundaries, Unicode normalization and case folding. The query
    /// accepts words, <c>"quoted phrases"</c>, <c>prefix*</c>, <c>AND</c>, <c>OR</c> and parentheses.
    /// </summary>
    Ranked
}

/// <summary>How adjacent ranked-query words without an explicit operator combine.</summary>
public enum OfficeDocumentQueryOperator {
    /// <summary>Every word must match (the default).</summary>
    And,
    /// <summary>Any word may match; blocks matching more words rank higher.</summary>
    Or
}

/// <summary>Ranking, query and resource settings for ranked document and chunk search.</summary>
public sealed class OfficeDocumentRankingOptions {
    /// <summary>BM25 term-frequency saturation, zero or more. Default: 1.2.</summary>
    public double K1 { get; set; } = 1.2;

    /// <summary>BM25 length normalization, from 0 to 1. Default: 0.75.</summary>
    public double B { get; set; } = 0.75;

    /// <summary>Weight of words found in the heading path. Default: 2.</summary>
    public double HeadingWeight { get; set; } = 2d;

    /// <summary>Weight of words found in the block or chunk text. Default: 1.</summary>
    public double BodyWeight { get; set; } = 1d;

    /// <summary>When true, "resume" also matches "résumé". Default: false.</summary>
    public bool AccentFolding { get; set; }

    /// <summary>Operator between adjacent words. Default: <see cref="OfficeDocumentQueryOperator.And"/>.</summary>
    public OfficeDocumentQueryOperator DefaultOperator { get; set; } = OfficeDocumentQueryOperator.And;

    /// <summary>Maximum word, phrase and prefix clauses in one query. Default: 64.</summary>
    public int MaximumClauses { get; set; } = 64;

    /// <summary>Maximum index terms one prefix expands to, in ordinal order. Default: 128.</summary>
    public int MaximumPrefixExpansions { get; set; } = 128;

    /// <summary>Maximum postings read for one query. Default: 10,000,000.</summary>
    public long MaximumPostingsVisited { get; set; } = 10_000_000;

    /// <summary>Maximum tokens indexed across the searched text; text beyond it is not searchable. Default: 20,000,000.</summary>
    public long MaximumIndexedTokens { get; set; } = 20_000_000;
}

/// <summary>One matched range in a block or chunk text.</summary>
public readonly struct OfficeDocumentSearchHighlight : IEquatable<OfficeDocumentSearchHighlight> {
    /// <summary>Creates a highlight.</summary>
    public OfficeDocumentSearchHighlight(int start, int length) {
        Start = start;
        Length = length;
    }

    /// <summary>Zero-based UTF-16 offset.</summary>
    public int Start { get; }

    /// <summary>Length in UTF-16 code units.</summary>
    public int Length { get; }

    /// <inheritdoc />
    public bool Equals(OfficeDocumentSearchHighlight other) => Start == other.Start && Length == other.Length;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeDocumentSearchHighlight other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => unchecked(Start * 397 ^ Length);

    /// <inheritdoc />
    public override string ToString() => Start + "+" + Length;
}
