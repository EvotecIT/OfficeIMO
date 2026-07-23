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
        OfficeDocumentSource source,
        string query,
        int totalPageCount,
        IReadOnlyList<OfficeDocumentSearchHit> hits,
        bool maximumResultsReached) {
        Source = source;
        Query = query;
        TotalPageCount = totalPageCount;
        Hits = hits;
        MaximumResultsReached = maximumResultsReached;
        PageNumbers = hits
            .SelectMany(static hit => hit.Pages)
            .Where(static location => location.Number.HasValue)
            .Select(static location => location.Number!.Value)
            .Distinct()
            .OrderBy(static number => number)
            .ToArray();
    }

    /// <summary>Source document that was searched.</summary>
    public OfficeDocumentSource Source { get; }

    /// <summary>Original query text.</summary>
    public string Query { get; }

    /// <summary>Total physical page count known for the read operation.</summary>
    public int TotalPageCount { get; }

    /// <summary>Matching occurrences in normalized source order.</summary>
    public IReadOnlyList<OfficeDocumentSearchHit> Hits { get; }

    /// <summary>Distinct one-based physical pages containing matches.</summary>
    public IReadOnlyList<int> PageNumbers { get; }

    /// <summary>
    /// True when the configured maximum result count was reached and the search stopped at that ceiling.
    /// </summary>
    public bool MaximumResultsReached { get; }
}

public static partial class OfficeDocumentReadResultExtensions {
    private const int MaximumFallbackCorrelationCharacters = 1024 * 1024;

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
            int remainingResultCount = effective.MaximumResults - hits.Count;
            IReadOnlyList<int> occurrenceIndexes = FindOccurrences(
                text,
                query,
                comparison,
                effective.WholeWord,
                remainingResultCount);
            if (occurrenceIndexes.Count == 0) {
                continue;
            }

            int returnedOccurrenceCount = occurrenceIndexes.Count;
            IReadOnlyList<OfficeDocumentPageLocation> locations = document.Locate(block);
            IReadOnlyList<IReadOnlyList<OfficeDocumentPageLocation>>? occurrencePages =
                CorrelateOccurrencesToPages(
                    locations,
                    block.Id,
                    text,
                    query,
                    comparison,
                    effective.WholeWord,
                    occurrenceIndexes,
                    returnedOccurrenceCount,
                    out bool hasConflictingPageEvidence);
            OfficeDocumentPageLocation[] pageSpecific =
                occurrencePages == null && !hasConflictingPageEvidence
                    ? locations
                        .Where(location => PageContainsQuery(
                            location.Page,
                            block.Id,
                            query,
                            comparison,
                            effective.WholeWord))
                        .ToArray()
                    : Array.Empty<OfficeDocumentPageLocation>();
            IReadOnlyList<OfficeDocumentPageLocation> pageFallback =
                Array.Empty<OfficeDocumentPageLocation>();
            if (occurrencePages == null && !hasConflictingPageEvidence) {
                bool hasPageBlockEvidence = locations.Any(location =>
                    (location.Page.Blocks ?? Array.Empty<OfficeDocumentBlock>())
                    .Any(pageBlock => string.Equals(pageBlock.Id, block.Id, StringComparison.Ordinal)));
                pageFallback = pageSpecific.Length > 0 || hasPageBlockEvidence
                    ? pageSpecific
                    : locations;
            }

            for (int occurrenceIndex = 0; occurrenceIndex < returnedOccurrenceCount; occurrenceIndex++) {
                hits.Add(new OfficeDocumentSearchHit(
                    block,
                    occurrenceIndexes[occurrenceIndex],
                    query.Length,
                    occurrencePages != null
                        ? occurrencePages[occurrenceIndex]
                        : pageFallback));
                if (hits.Count >= effective.MaximumResults) {
                    return new OfficeDocumentSearchResult(
                        document.Source ?? new OfficeDocumentSource(),
                        query,
                        document.GetTotalPageCount(),
                        hits.AsReadOnly(),
                        maximumResultsReached: true);
                }
            }
        }

        return new OfficeDocumentSearchResult(
            document.Source ?? new OfficeDocumentSource(),
            query,
            document.GetTotalPageCount(),
            hits.AsReadOnly(),
            maximumResultsReached: false);
    }

    private static IReadOnlyList<IReadOnlyList<OfficeDocumentPageLocation>>? CorrelateOccurrencesToPages(
        IReadOnlyList<OfficeDocumentPageLocation> locations,
        string blockId,
        string sourceText,
        string query,
        StringComparison comparison,
        bool wholeWord,
        IReadOnlyList<int> sourceOccurrenceIndexes,
        int returnedOccurrenceCount,
        out bool hasConflictingPageEvidence) {
        hasConflictingPageEvidence = false;
        if (returnedOccurrenceCount == 0) {
            return Array.Empty<IReadOnlyList<OfficeDocumentPageLocation>>();
        }

        IReadOnlyList<IReadOnlyList<OfficeDocumentPageLocation>>? exactLocations =
            MapOccurrencesByFragmentOffsets(
                locations,
                blockId,
                sourceText,
                query.Length,
                sourceOccurrenceIndexes,
                returnedOccurrenceCount);
        if (exactLocations != null) {
            return exactLocations;
        }

        return MapOccurrencesByStableOrder(
            locations,
            blockId,
            sourceText,
            query,
            comparison,
            wholeWord,
            sourceOccurrenceIndexes,
            out hasConflictingPageEvidence);
    }

    private static IReadOnlyList<IReadOnlyList<OfficeDocumentPageLocation>>?
        MapOccurrencesByStableOrder(
            IReadOnlyList<OfficeDocumentPageLocation> locations,
            string blockId,
            string sourceText,
            string query,
            StringComparison comparison,
            bool wholeWord,
            IReadOnlyList<int> sourceOccurrenceIndexes,
            out bool hasConflictingPageEvidence) {
        hasConflictingPageEvidence = false;
        if (sourceText.Length > MaximumFallbackCorrelationCharacters) {
            hasConflictingPageEvidence = true;
            return null;
        }
        if (HasRepeatedCompleteBlockCopies(locations, blockId, sourceText)) {
            return null;
        }
        if (!TryGetCaseInsensitiveOccurrenceOrdinals(
                sourceText,
                query,
                sourceOccurrenceIndexes,
                out IReadOnlyList<int> stableOrdinals)) {
            hasConflictingPageEvidence = true;
            return null;
        }

        var combined = new StringBuilder();
        var fragmentStarts = new List<int>();
        var fragmentEnds = new List<int>();
        var fragmentLocations = new List<OfficeDocumentPageLocation>();
        foreach (OfficeDocumentPageLocation location in locations) {
            foreach (OfficeDocumentBlock pageBlock in location.Page.Blocks ?? Array.Empty<OfficeDocumentBlock>()) {
                if (!string.Equals(pageBlock.Id, blockId, StringComparison.Ordinal)) {
                    continue;
                }

                string fragmentText = pageBlock.Text ?? string.Empty;
                if (fragmentText.Length > MaximumFallbackCorrelationCharacters - combined.Length) {
                    hasConflictingPageEvidence = true;
                    return null;
                }
                fragmentStarts.Add(combined.Length);
                combined.Append(fragmentText);
                fragmentEnds.Add(combined.Length);
                fragmentLocations.Add(location);
            }
        }

        string combinedText = combined.ToString();
        if (CountCaseInsensitiveOccurrences(combinedText, query) !=
            CountCaseInsensitiveOccurrences(sourceText, query)) {
            hasConflictingPageEvidence = true;
            return null;
        }
        var fragmentOccurrenceIndexes = new List<int>(stableOrdinals.Count);
        bool foundSelectedOccurrences = CollectOccurrenceIndexesAtOrdinals(
            combinedText,
            query,
            stableOrdinals,
            fragmentOccurrenceIndexes);
        if (!foundSelectedOccurrences) {
            hasConflictingPageEvidence = true;
            return null;
        }

        foreach (int occurrenceIndex in fragmentOccurrenceIndexes) {
            if (string.Compare(
                    combinedText,
                    occurrenceIndex,
                    query,
                    0,
                    query.Length,
                    comparison) != 0 ||
                (wholeWord && !IsWholeWord(combinedText, occurrenceIndex, query.Length))) {
                hasConflictingPageEvidence = true;
                return null;
            }
        }

        return MapOccurrenceRangesToPages(
            fragmentStarts,
            fragmentEnds,
            fragmentLocations,
            fragmentOccurrenceIndexes,
            query.Length,
            fragmentOccurrenceIndexes.Count);
    }

    private static bool TryGetCaseInsensitiveOccurrenceOrdinals(
        string text,
        string query,
        IReadOnlyList<int> selectedOccurrenceIndexes,
        out IReadOnlyList<int> selectedOrdinals) {
        var ordinals = new List<int>(selectedOccurrenceIndexes.Count);
        int selectedIndex = 0;
        int ordinal = 0;
        int searchFrom = 0;
        while (selectedIndex < selectedOccurrenceIndexes.Count &&
               searchFrom <= text.Length - query.Length) {
            int occurrenceIndex = text.IndexOf(query, searchFrom, StringComparison.OrdinalIgnoreCase);
            if (occurrenceIndex < 0 || occurrenceIndex > selectedOccurrenceIndexes[selectedIndex]) {
                selectedOrdinals = null!;
                return false;
            }
            if (occurrenceIndex == selectedOccurrenceIndexes[selectedIndex]) {
                ordinals.Add(ordinal);
                selectedIndex++;
            }
            ordinal++;
            searchFrom = occurrenceIndex + Math.Max(1, query.Length);
        }

        if (selectedIndex != selectedOccurrenceIndexes.Count) {
            selectedOrdinals = null!;
            return false;
        }
        selectedOrdinals = ordinals.AsReadOnly();
        return true;
    }

    private static bool HasRepeatedCompleteBlockCopies(
        IReadOnlyList<OfficeDocumentPageLocation> locations,
        string blockId,
        string sourceText) {
        int completeCopyCount = 0;
        foreach (OfficeDocumentPageLocation location in locations) {
            foreach (OfficeDocumentBlock pageBlock in location.Page.Blocks ?? Array.Empty<OfficeDocumentBlock>()) {
                if (string.Equals(pageBlock.Id, blockId, StringComparison.Ordinal) &&
                    string.Equals(pageBlock.Text, sourceText, StringComparison.Ordinal) &&
                    ++completeCopyCount > 1) {
                    return true;
                }
            }
        }
        return false;
    }

    private static bool CollectOccurrenceIndexesAtOrdinals(
        string text,
        string query,
        IReadOnlyList<int> selectedOrdinals,
        List<int> selectedOccurrenceIndexes) {
        if (selectedOrdinals.Count == 0) return true;
        int selectedOrdinalIndex = 0;
        int searchFrom = 0;
        int totalOccurrenceCount = 0;
        while (searchFrom <= text.Length - query.Length) {
            int index = text.IndexOf(query, searchFrom, StringComparison.OrdinalIgnoreCase);
            if (index < 0) {
                break;
            }

            searchFrom = index + Math.Max(1, query.Length);
            if (selectedOrdinalIndex < selectedOrdinals.Count &&
                totalOccurrenceCount == selectedOrdinals[selectedOrdinalIndex]) {
                selectedOccurrenceIndexes.Add(index);
                selectedOrdinalIndex++;
                if (selectedOrdinalIndex == selectedOrdinals.Count) {
                    return true;
                }
            }
            totalOccurrenceCount++;
        }
        return false;
    }

    private static int CountCaseInsensitiveOccurrences(string text, string query) {
        int count = 0;
        int searchFrom = 0;
        while (searchFrom <= text.Length - query.Length) {
            int index = text.IndexOf(query, searchFrom, StringComparison.OrdinalIgnoreCase);
            if (index < 0) break;
            count++;
            searchFrom = index + Math.Max(1, query.Length);
        }
        return count;
    }

    private static IReadOnlyList<IReadOnlyList<OfficeDocumentPageLocation>>?
        MapOccurrencesByFragmentOffsets(
            IReadOnlyList<OfficeDocumentPageLocation> locations,
            string blockId,
            string sourceText,
            int queryLength,
            IReadOnlyList<int> sourceOccurrenceIndexes,
            int occurrenceCount) {
        var fragmentStarts = new List<int>();
        var fragmentEnds = new List<int>();
        var fragmentLocations = new List<OfficeDocumentPageLocation>();
        int combinedLength = 0;
        bool fragmentsMatchSource = true;
        foreach (OfficeDocumentPageLocation location in locations) {
            foreach (OfficeDocumentBlock pageBlock in location.Page.Blocks ?? Array.Empty<OfficeDocumentBlock>()) {
                if (!string.Equals(pageBlock.Id, blockId, StringComparison.Ordinal)) {
                    continue;
                }

                string fragmentText = pageBlock.Text ?? string.Empty;
                fragmentStarts.Add(combinedLength);
                if (fragmentsMatchSource &&
                    (fragmentText.Length > sourceText.Length - combinedLength ||
                     string.CompareOrdinal(
                         sourceText,
                         combinedLength,
                         fragmentText,
                         0,
                         fragmentText.Length) != 0)) {
                    fragmentsMatchSource = false;
                }
                combinedLength += fragmentText.Length;
                fragmentEnds.Add(combinedLength);
                fragmentLocations.Add(location);
            }
        }

        if (!fragmentsMatchSource || combinedLength != sourceText.Length) {
            return null;
        }

        return MapOccurrenceRangesToPages(
            fragmentStarts,
            fragmentEnds,
            fragmentLocations,
            sourceOccurrenceIndexes,
            queryLength,
            occurrenceCount);
    }

    private static IReadOnlyList<IReadOnlyList<OfficeDocumentPageLocation>>?
        MapOccurrenceRangesToPages(
            IReadOnlyList<int> fragmentStarts,
            IReadOnlyList<int> fragmentEnds,
            IReadOnlyList<OfficeDocumentPageLocation> fragmentLocations,
            IReadOnlyList<int> occurrenceIndexes,
            int queryLength,
            int occurrenceCount) {
        var mapped =
            new List<IReadOnlyList<OfficeDocumentPageLocation>>(occurrenceCount);
        int firstPossibleFragment = 0;
        for (int occurrenceIndex = 0; occurrenceIndex < occurrenceCount; occurrenceIndex++) {
            int occurrenceStart = occurrenceIndexes[occurrenceIndex];
            int occurrenceEnd = occurrenceStart + queryLength;
            while (firstPossibleFragment < fragmentLocations.Count
                   && fragmentEnds[firstPossibleFragment] <= occurrenceStart) {
                firstPossibleFragment++;
            }
            var occurrenceLocations = new List<OfficeDocumentPageLocation>();
            for (int fragmentIndex = firstPossibleFragment;
                 fragmentIndex < fragmentLocations.Count
                 && fragmentStarts[fragmentIndex] < occurrenceEnd;
                 fragmentIndex++) {
                if (fragmentEnds[fragmentIndex] <= occurrenceStart) continue;

                OfficeDocumentPageLocation location = fragmentLocations[fragmentIndex];
                if (!occurrenceLocations.Contains(location)) {
                    occurrenceLocations.Add(location);
                }
            }

            if (occurrenceLocations.Count == 0) {
                return null;
            }
            mapped.Add(occurrenceLocations.AsReadOnly());
        }
        return mapped.AsReadOnly();
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

            if (CountOccurrences(
                    pageBlock.Text ?? string.Empty,
                    query,
                    comparison,
                    wholeWord,
                    maximumOccurrences: 1) > 0) {
                return true;
            }
        }
        return false;
    }

    private static IReadOnlyList<int> FindOccurrences(
        string text,
        string query,
        StringComparison comparison,
        bool wholeWord,
        int maximumOccurrences) {
        var occurrences = new List<int>();
        if (maximumOccurrences < 1) {
            return occurrences.AsReadOnly();
        }

        int searchFrom = 0;
        while (searchFrom <= text.Length - query.Length) {
            int index = text.IndexOf(query, searchFrom, comparison);
            if (index < 0) {
                break;
            }

            searchFrom = index + Math.Max(1, query.Length);
            if (!wholeWord || IsWholeWord(text, index, query.Length)) {
                occurrences.Add(index);
                if (occurrences.Count >= maximumOccurrences) {
                    break;
                }
            }
        }
        return occurrences.AsReadOnly();
    }

    private static int CountOccurrences(
        string text,
        string query,
        StringComparison comparison,
        bool wholeWord,
        int maximumOccurrences) {
        if (maximumOccurrences < 1) {
            return 0;
        }

        int count = 0;
        int searchFrom = 0;
        while (searchFrom <= text.Length - query.Length) {
            int index = text.IndexOf(query, searchFrom, comparison);
            if (index < 0) {
                break;
            }

            searchFrom = index + Math.Max(1, query.Length);
            if (!wholeWord || IsWholeWord(text, index, query.Length)) {
                count++;
                if (count >= maximumOccurrences) {
                    break;
                }
            }
        }
        return count;
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
