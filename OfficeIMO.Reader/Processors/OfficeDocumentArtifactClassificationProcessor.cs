using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Reader;

/// <summary>Options for repeated page-boundary artifact classification.</summary>
public sealed class OfficeDocumentArtifactClassificationOptions {
    /// <summary>Minimum number of distinct pages on which boundary text must repeat.</summary>
    public int MinimumPageOccurrences { get; set; } = 2;

    /// <summary>Number of non-empty blocks inspected at each page boundary.</summary>
    public int BoundaryBlockCount { get; set; } = 2;

    /// <summary>Maximum normalized text length eligible for artifact classification.</summary>
    public int MaximumTextLength { get; set; } = 200;

    internal OfficeDocumentArtifactClassificationOptions Clone() => new OfficeDocumentArtifactClassificationOptions {
        MinimumPageOccurrences = MinimumPageOccurrences,
        BoundaryBlockCount = BoundaryBlockCount,
        MaximumTextLength = MaximumTextLength
    };
}

/// <summary>
/// Classifies short text repeated at page starts or ends as header, footer, or artifact blocks.
/// </summary>
public sealed class OfficeDocumentArtifactClassificationProcessor : OfficeDocumentProcessorBase {
    private readonly OfficeDocumentArtifactClassificationOptions _options;

    /// <summary>Creates the processor.</summary>
    public OfficeDocumentArtifactClassificationProcessor(OfficeDocumentArtifactClassificationOptions? options = null)
        : base("officeimo.reader.classify-artifacts") {
        _options = (options ?? new OfficeDocumentArtifactClassificationOptions()).Clone();
        if (_options.MinimumPageOccurrences < 2) throw new ArgumentOutOfRangeException(nameof(options), "Minimum page occurrences must be at least 2.");
        if (_options.BoundaryBlockCount < 1) throw new ArgumentOutOfRangeException(nameof(options), "Boundary block count must be positive.");
        if (_options.MaximumTextLength < 1) throw new ArgumentOutOfRangeException(nameof(options), "Maximum text length must be positive.");
    }

    /// <inheritdoc />
    public override OfficeDocumentReadResult Process(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        IReadOnlyList<OfficeDocumentPage> pages = document.Pages ?? Array.Empty<OfficeDocumentPage>();
        if (pages.Count < _options.MinimumPageOccurrences) return document;

        Dictionary<string, int> headerCounts = CountBoundaryText(pages, fromStart: true, context);
        Dictionary<string, int> footerCounts = CountBoundaryText(pages, fromStart: false, context);
        HashSet<string> headers = SelectRepeated(headerCounts);
        HashSet<string> footers = SelectRepeated(footerCounts);
        if (headers.Count == 0 && footers.Count == 0) return document;

        ApplyBoundaryClassifications(document, pages, headers, footers, context);
        return document;
    }

    private void ApplyBoundaryClassifications(
        OfficeDocumentReadResult document,
        IReadOnlyList<OfficeDocumentPage> pages,
        ISet<string> headers,
        ISet<string> footers,
        OfficeDocumentProcessorContext context) {
        var classifications = new Dictionary<OfficeDocumentBlock, string>(ReferenceIdentityComparer<OfficeDocumentBlock>.Instance);
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            context.CancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<OfficeDocumentBlock> blocks = pages[pageIndex]?.Blocks ?? Array.Empty<OfficeDocumentBlock>();
            ClassifyBoundarySide(blocks, fromStart: true, headers, "header", classifications);
            ClassifyBoundarySide(blocks, fromStart: false, footers, "footer", classifications);
        }

        var classificationsById = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (KeyValuePair<OfficeDocumentBlock, string> classification in classifications) {
            classification.Key.Kind = classification.Value;
            if (!string.IsNullOrWhiteSpace(classification.Key.Id)) {
                MergeClassification(classificationsById, classification.Key.Id, classification.Value);
            }
        }

        foreach (OfficeDocumentBlock block in document.Blocks ?? Array.Empty<OfficeDocumentBlock>()) {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (block != null &&
                !string.IsNullOrWhiteSpace(block.Id) &&
                classificationsById.TryGetValue(block.Id, out string? kind)) {
                block.Kind = kind;
            }
        }
    }

    private void ClassifyBoundarySide(
        IReadOnlyList<OfficeDocumentBlock> blocks,
        bool fromStart,
        ISet<string> repeated,
        string kind,
        IDictionary<OfficeDocumentBlock, string> classifications) {
        int accepted = 0;
        for (int offset = 0; offset < blocks.Count && accepted < _options.BoundaryBlockCount; offset++) {
            int blockIndex = fromStart ? offset : blocks.Count - 1 - offset;
            OfficeDocumentBlock? block = blocks[blockIndex];
            string text = NormalizeText(block?.Text);
            if (text.Length == 0 || text.Length > _options.MaximumTextLength) continue;
            accepted++;
            if (block != null && repeated.Contains(text)) MergeClassification(classifications, block, kind);
        }
    }

    private static void MergeClassification<TKey>(IDictionary<TKey, string> classifications, TKey key, string kind) {
        if (classifications.TryGetValue(key, out string? existing) && !string.Equals(existing, kind, StringComparison.Ordinal)) {
            classifications[key] = "artifact";
        } else {
            classifications[key] = kind;
        }
    }

    private Dictionary<string, int> CountBoundaryText(
        IReadOnlyList<OfficeDocumentPage> pages,
        bool fromStart,
        OfficeDocumentProcessorContext context) {
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            context.CancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<OfficeDocumentBlock> blocks = pages[pageIndex]?.Blocks ?? Array.Empty<OfficeDocumentBlock>();
            var onPage = new HashSet<string>(StringComparer.Ordinal);
            int accepted = 0;
            for (int offset = 0; offset < blocks.Count && accepted < _options.BoundaryBlockCount; offset++) {
                int blockIndex = fromStart ? offset : blocks.Count - 1 - offset;
                string text = NormalizeText(blocks[blockIndex]?.Text);
                if (text.Length == 0 || text.Length > _options.MaximumTextLength) continue;
                accepted++;
                onPage.Add(text);
            }
            foreach (string text in onPage) {
                counts[text] = counts.TryGetValue(text, out int count) ? count + 1 : 1;
            }
        }
        return counts;
    }

    private HashSet<string> SelectRepeated(Dictionary<string, int> counts) {
        var repeated = new HashSet<string>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, int> entry in counts) {
            if (entry.Value >= _options.MinimumPageOccurrences) repeated.Add(entry.Key);
        }
        return repeated;
    }

    private static string NormalizeText(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var normalized = new StringBuilder(value!.Length);
        bool pendingSpace = false;
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (char.IsWhiteSpace(character)) {
                pendingSpace = normalized.Length > 0;
                continue;
            }
            if (pendingSpace) normalized.Append(' ');
            normalized.Append(character);
            pendingSpace = false;
        }
        return normalized.ToString();
    }
}
