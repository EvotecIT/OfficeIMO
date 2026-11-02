using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Text.RegularExpressions;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentReadResultExtensions {
    private static OfficeDocumentSearchResult SearchRegularExpression(OfficeDocumentReadResult document, string pattern,
        OfficeDocumentSearchOptions options, CancellationToken cancellationToken) {
        if (options.RegularExpressionTimeout <= TimeSpan.Zero || options.RegularExpressionTimeout == Regex.InfiniteMatchTimeout)
            throw new ArgumentOutOfRangeException(nameof(options), "Regular-expression timeout must be positive and finite.");
        var flags = RegexOptions.CultureInvariant | (options.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase);
        var expression = new Regex(pattern, flags, options.RegularExpressionTimeout);
        var hits = new List<OfficeDocumentSearchHit>();
        var comparison = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        foreach (OfficeDocumentBlock block in document.Blocks ?? Array.Empty<OfficeDocumentBlock>()) {
            cancellationToken.ThrowIfCancellationRequested();
            string text = block.Text ?? string.Empty;
            IReadOnlyList<OfficeDocumentPageLocation> locations = document.Locate(block);
            bool exactFragments = TryGetSourceFragmentRanges(locations, block.Id, text, out var starts, out var ends, out var fragmentLocations);
            var matches = new List<Match>();
            foreach (Match match in expression.Matches(text)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (options.WholeWord && !IsWholeWord(text, match.Index, match.Length)) continue;
                matches.Add(match);
                if (matches.Count >= options.MaximumResults - hits.Count) break;
            }
            var correlatedPages = !exactFragments && locations.Count > 0
                ? CorrelateRegularExpressionMatches(locations, block.Id, text, matches, expression, comparison, options.WholeWord, cancellationToken)
                : null;
            int firstFragment = 0;
            foreach (Match match in matches) {
                cancellationToken.ThrowIfCancellationRequested();
                int pageIndex = match.Length == 0 && text.Length > 0 ? Math.Min(match.Index, text.Length - 1) : match.Index;
                int pageLength = match.Length == 0 && text.Length > 0 ? 1 : match.Length;
                IReadOnlyList<OfficeDocumentPageLocation> pages;
                if (locations.Count == 0) {
                    pages = Array.Empty<OfficeDocumentPageLocation>();
                } else if (exactFragments && text.Length > 0) {
                    while (firstFragment < ends.Count && ends[firstFragment] <= pageIndex) firstFragment++;
                    var matchedPages = new List<OfficeDocumentPageLocation>();
                    for (int i = firstFragment; i < starts.Count && starts[i] < pageIndex + pageLength; i++) {
                        if (ends[i] > pageIndex && !matchedPages.Contains(fragmentLocations[i])) matchedPages.Add(fragmentLocations[i]);
                    }
                    pages = matchedPages;
                } else if (match.Length > 0) {
                    pages = correlatedPages![match.Index];
                } else {
                    pages = text.Length == 0 ? locations : Array.Empty<OfficeDocumentPageLocation>();
                }
                hits.Add(new OfficeDocumentSearchHit(block, match.Index, match.Length, pages));
                if (hits.Count >= options.MaximumResults) return Result(true);
            }
        }
        return Result(false);

        OfficeDocumentSearchResult Result(bool limited) => new(document.Source ?? new OfficeDocumentSource(), pattern,
            document.GetTotalPageCount(), hits.AsReadOnly(), limited);
    }

    private static Dictionary<int, IReadOnlyList<OfficeDocumentPageLocation>> CorrelateRegularExpressionMatches(
        IReadOnlyList<OfficeDocumentPageLocation> locations, string blockId, string text, IReadOnlyList<Match> matches,
        Regex expression, StringComparison comparison, bool wholeWord, CancellationToken cancellationToken) {
        var unknown = matches.Where(match => match.Length > 0).ToDictionary(match => match.Index,
            _ => (IReadOnlyList<OfficeDocumentPageLocation>)Array.Empty<OfficeDocumentPageLocation>());
        if (unknown.Count == 0 || text.Length > MaximumFallbackCorrelationCharacters) return unknown;
        cancellationToken.ThrowIfCancellationRequested();
        if (HasRepeatedCompleteBlockCopies(locations, blockId, text)) {
            return CorrelateRepeatedRegularExpressionBlocks(locations, blockId, matches, expression,
                comparison, wholeWord, cancellationToken);
        }

        // Build normalized producer evidence once. Applying the same expression once
        // to each text correlates variable values without a scan/allocation per value.
        var combined = new StringBuilder();
        var starts = new List<int>();
        var ends = new List<int>();
        var fragments = new List<OfficeDocumentPageLocation>();
        foreach (var location in locations) {
            foreach (var block in location.Page.Blocks ?? Array.Empty<OfficeDocumentBlock>()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.Equals(block.Id, blockId, StringComparison.Ordinal)) continue;
                string fragment = block.Text ?? string.Empty;
                if (fragment.Length > MaximumFallbackCorrelationCharacters - combined.Length) return unknown;
                starts.Add(combined.Length);
                combined.Append(fragment);
                ends.Add(combined.Length);
                fragments.Add(location);
            }
        }
        if (fragments.Count == 0) {
            foreach (var match in matches.Where(match => match.Length > 0)) unknown[match.Index] = locations;
            return unknown;
        }

        var pages = new Dictionary<int, IReadOnlyList<OfficeDocumentPageLocation>>();
        using var sourceMatches = EnumerateRegularExpressionMatches(expression, text, wholeWord, cancellationToken).GetEnumerator();
        using var fragmentMatches = EnumerateRegularExpressionMatches(expression, combined.ToString(), wholeWord, cancellationToken).GetEnumerator();
        int firstFragment = 0;
        while (sourceMatches.MoveNext()) {
            if (!fragmentMatches.MoveNext() || !string.Equals(sourceMatches.Current.Value, fragmentMatches.Current.Value, comparison))
                return unknown;
            Match source = sourceMatches.Current;
            if (!unknown.ContainsKey(source.Index)) continue;
            Match fragment = fragmentMatches.Current;
            while (firstFragment < ends.Count && ends[firstFragment] <= fragment.Index) firstFragment++;
            var containing = new List<OfficeDocumentPageLocation>();
            for (int i = firstFragment; i < starts.Count && starts[i] < fragment.Index + fragment.Length; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (ends[i] > fragment.Index && !containing.Contains(fragments[i])) containing.Add(fragments[i]);
            }
            pages.Add(source.Index, containing);
        }
        return fragmentMatches.MoveNext() ? unknown : pages;
    }

    private static Dictionary<int, IReadOnlyList<OfficeDocumentPageLocation>> CorrelateRepeatedRegularExpressionBlocks(
        IReadOnlyList<OfficeDocumentPageLocation> locations, string blockId, IReadOnlyList<Match> matches, Regex expression,
        StringComparison comparison, bool wholeWord, CancellationToken cancellationToken) {
        var comparer = comparison == StringComparison.Ordinal ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
        var byValue = new Dictionary<string, List<OfficeDocumentPageLocation>>(comparer);
        int characters = 0;
        foreach (var location in locations) {
            foreach (var block in location.Page.Blocks ?? Array.Empty<OfficeDocumentBlock>()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.Equals(block.Id, blockId, StringComparison.Ordinal)) continue;
                string fragment = block.Text ?? string.Empty;
                if (fragment.Length > MaximumFallbackCorrelationCharacters - characters)
                    return matches.Where(match => match.Length > 0).ToDictionary(match => match.Index,
                        _ => (IReadOnlyList<OfficeDocumentPageLocation>)Array.Empty<OfficeDocumentPageLocation>());
                characters += fragment.Length;
                foreach (Match match in EnumerateRegularExpressionMatches(expression, fragment, wholeWord, cancellationToken)) {
                    if (match.Length == 0) continue;
                    if (!byValue.TryGetValue(match.Value, out var pages)) byValue.Add(match.Value, pages = new List<OfficeDocumentPageLocation>());
                    if (pages.Count == 0 || !ReferenceEquals(pages[pages.Count - 1], location)) pages.Add(location);
                }
            }
        }
        return matches.Where(match => match.Length > 0).ToDictionary(match => match.Index,
            match => byValue.TryGetValue(match.Value, out var pages)
                ? (IReadOnlyList<OfficeDocumentPageLocation>)pages : Array.Empty<OfficeDocumentPageLocation>());
    }

    private static IEnumerable<Match> EnumerateRegularExpressionMatches(Regex expression, string text, bool wholeWord,
        CancellationToken cancellationToken) {
        foreach (Match match in expression.Matches(text)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!wholeWord || IsWholeWord(text, match.Index, match.Length)) yield return match;
        }
    }
}
