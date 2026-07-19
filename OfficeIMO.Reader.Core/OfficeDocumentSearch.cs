using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Reader;

/// <summary>Controls normalized block search in a document read result.</summary>
public sealed class OfficeDocumentSearchOptions {
    /// <summary>When true, uses case-sensitive ordinal matching.</summary>
    public bool MatchCase { get; set; }

    /// <summary>When true, accepts matches surrounded by non-word characters only.</summary>
    public bool WholeWord { get; set; }

    /// <summary>Maximum number of occurrences returned. Default: 1,000.</summary>
    public int MaximumResults { get; set; } = 1000;
}

/// <summary>One query occurrence and its source/page locations.</summary>
public sealed class OfficeDocumentSearchHit {
    internal OfficeDocumentSearchHit(
        OfficeDocumentBlock block,
        int startIndex,
        int length,
        IReadOnlyList<OfficeDocumentPageLocation> pages) {
        Block = block;
        StartIndex = startIndex;
        Length = length;
        Pages = pages;
    }

    /// <summary>Normalized source block containing the occurrence.</summary>
    public OfficeDocumentBlock Block { get; }

    /// <summary>Zero-based occurrence offset in <see cref="OfficeDocumentBlock.Text"/>.</summary>
    public int StartIndex { get; }

    /// <summary>Occurrence length.</summary>
    public int Length { get; }

    /// <summary>Page-like locations that contain the matching source block fragment.</summary>
    public IReadOnlyList<OfficeDocumentPageLocation> Pages { get; }
}

/// <summary>Search results with aggregated page citation information.</summary>
public sealed class OfficeDocumentSearchResult {
    internal OfficeDocumentSearchResult(
        string query,
        int totalPageCount,
        IReadOnlyList<OfficeDocumentSearchHit> hits) {
        Query = query;
        TotalPageCount = totalPageCount;
        Hits = hits;
        PageNumbers = hits
            .SelectMany(static hit => hit.Pages)
            .Where(static location => location.Number.HasValue)
            .Select(static location => location.Number!.Value)
            .Distinct()
            .OrderBy(static number => number)
            .ToArray();
    }

    /// <summary>Original query text.</summary>
    public string Query { get; }

    /// <summary>Total physical page count known for the read operation.</summary>
    public int TotalPageCount { get; }

    /// <summary>Matching occurrences in normalized source order.</summary>
    public IReadOnlyList<OfficeDocumentSearchHit> Hits { get; }

    /// <summary>Distinct one-based physical pages containing matches.</summary>
    public IReadOnlyList<int> PageNumbers { get; }
}

public static partial class OfficeDocumentReadResultExtensions {
    /// <summary>
    /// Searches normalized document blocks and attaches page-like locations to every occurrence.
    /// </summary>
    public static OfficeDocumentSearchResult Search(
        this OfficeDocumentReadResult document,
        string query,
        OfficeDocumentSearchOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrEmpty(query)) throw new ArgumentException("Search query cannot be empty.", nameof(query));

        OfficeDocumentSearchOptions effective = options ?? new OfficeDocumentSearchOptions();
        if (effective.MaximumResults < 1) {
            throw new ArgumentOutOfRangeException(nameof(options), "MaximumResults must be positive.");
        }

        StringComparison comparison = effective.MatchCase
            ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;
        var hits = new List<OfficeDocumentSearchHit>();

        foreach (OfficeDocumentBlock block in document.Blocks ?? Array.Empty<OfficeDocumentBlock>()) {
            string text = block.Text ?? string.Empty;
            int searchFrom = 0;
            while (searchFrom <= text.Length - query.Length) {
                int index = text.IndexOf(query, searchFrom, comparison);
                if (index < 0) {
                    break;
                }

                searchFrom = index + Math.Max(1, query.Length);
                if (effective.WholeWord && !IsWholeWord(text, index, query.Length)) {
                    continue;
                }

                IReadOnlyList<OfficeDocumentPageLocation> locations = document.Locate(block);
                OfficeDocumentPageLocation[] pageSpecific = locations
                    .Where(location => PageContainsQuery(
                        location.Page,
                        block.Id,
                        query,
                        comparison,
                        effective.WholeWord))
                    .ToArray();
                bool hasPageBlockEvidence = locations.Any(location =>
                    (location.Page.Blocks ?? Array.Empty<OfficeDocumentBlock>())
                    .Any(pageBlock => string.Equals(pageBlock.Id, block.Id, StringComparison.Ordinal)));
                hits.Add(new OfficeDocumentSearchHit(
                    block,
                    index,
                    query.Length,
                    pageSpecific.Length > 0 || hasPageBlockEvidence
                        ? pageSpecific
                        : locations));
                if (hits.Count >= effective.MaximumResults) {
                    return new OfficeDocumentSearchResult(query, document.GetTotalPageCount(), hits.AsReadOnly());
                }
            }
        }

        return new OfficeDocumentSearchResult(query, document.GetTotalPageCount(), hits.AsReadOnly());
    }

    private static bool PageContainsQuery(
        OfficeDocumentPage page,
        string blockId,
        string query,
        StringComparison comparison,
        bool wholeWord) {
        foreach (OfficeDocumentBlock pageBlock in page.Blocks ?? Array.Empty<OfficeDocumentBlock>()) {
            if (!string.Equals(pageBlock.Id, blockId, StringComparison.Ordinal)) {
                continue;
            }

            string text = pageBlock.Text ?? string.Empty;
            int searchFrom = 0;
            while (searchFrom <= text.Length - query.Length) {
                int index = text.IndexOf(query, searchFrom, comparison);
                if (index < 0) {
                    break;
                }
                if (!wholeWord || IsWholeWord(text, index, query.Length)) {
                    return true;
                }
                searchFrom = index + Math.Max(1, query.Length);
            }
        }
        return false;
    }

    private static bool IsWholeWord(string text, int start, int length) {
        bool startsAtBoundary = start == 0 || !IsWordCharacter(text[start - 1]);
        int end = start + length;
        bool endsAtBoundary = end >= text.Length || !IsWordCharacter(text[end]);
        return startsAtBoundary && endsAtBoundary;
    }

    private static bool IsWordCharacter(char value) =>
        char.IsLetterOrDigit(value) || value == '_';
}
