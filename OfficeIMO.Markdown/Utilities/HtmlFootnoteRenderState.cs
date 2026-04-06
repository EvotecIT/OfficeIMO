namespace OfficeIMO.Markdown;

internal sealed class HtmlFootnoteRenderState {
    private readonly Dictionary<string, FootnoteDefinitionBlock> _definitionsByLabel;
    private readonly List<string> _orderedReferencedLabels = new();
    private readonly Dictionary<string, int> _numbersByLabel;
    private readonly Dictionary<string, List<string>> _referenceIdsByLabel;

    private HtmlFootnoteRenderState(Dictionary<string, FootnoteDefinitionBlock> definitionsByLabel) {
        _definitionsByLabel = definitionsByLabel;
        _numbersByLabel = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        _referenceIdsByLabel = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
    }

    internal IReadOnlyList<string> OrderedReferencedLabels => _orderedReferencedLabels;

    internal static HtmlFootnoteRenderState Create(IReadOnlyList<IMarkdownBlock> blocks) {
        var definitions = new Dictionary<string, FootnoteDefinitionBlock>(StringComparer.OrdinalIgnoreCase);
        if (blocks != null) {
            for (int i = 0; i < blocks.Count; i++) {
                if (blocks[i] is FootnoteDefinitionBlock footnote && !string.IsNullOrWhiteSpace(footnote.Label)) {
                    definitions[footnote.Label] = footnote;
                }
            }
        }

        return new HtmlFootnoteRenderState(definitions);
    }

    internal bool IsDefined(string label) =>
        !string.IsNullOrWhiteSpace(label) && _definitionsByLabel.ContainsKey(label);

    internal FootnoteDefinitionBlock? FindDefinition(string label) {
        if (string.IsNullOrWhiteSpace(label)) {
            return null;
        }

        return _definitionsByLabel.TryGetValue(label, out var definition)
            ? definition
            : null;
    }

    internal HtmlFootnoteReferenceInfo RegisterReference(string label) {
        if (!IsDefined(label)) {
            throw new InvalidOperationException($"Footnote '{label}' is not defined.");
        }

        if (!_numbersByLabel.TryGetValue(label, out var number)) {
            number = _orderedReferencedLabels.Count + 1;
            _numbersByLabel[label] = number;
            _orderedReferencedLabels.Add(label);
        }

        if (!_referenceIdsByLabel.TryGetValue(label, out var refIds)) {
            refIds = new List<string>();
            _referenceIdsByLabel[label] = refIds;
        }

        var suffix = refIds.Count == 0 ? string.Empty : "-" + (refIds.Count + 1).ToString();
        var escaped = EscapeFootnoteIdFragment(label);
        var referenceId = "fnref-" + escaped + suffix;
        refIds.Add(referenceId);

        return new HtmlFootnoteReferenceInfo(label, escaped, number, referenceId, refIds.Count);
    }

    internal HtmlFootnoteDefinitionInfo? GetDefinitionInfo(string label) {
        if (!IsDefined(label) || !_numbersByLabel.TryGetValue(label, out var number)) {
            return null;
        }

        var escaped = EscapeFootnoteIdFragment(label);
        var refIds = _referenceIdsByLabel.TryGetValue(label, out var refs)
            ? (IReadOnlyList<string>)refs
            : Array.Empty<string>();

        return new HtmlFootnoteDefinitionInfo(label, escaped, number, refIds);
    }

    internal static string EscapeFootnoteIdFragment(string label) =>
        Uri.EscapeDataString(label ?? string.Empty);
}

internal struct HtmlFootnoteReferenceInfo {
    internal HtmlFootnoteReferenceInfo(
        string label,
        string escapedLabel,
        int number,
        string referenceId,
        int referenceIndex) {
        Label = label;
        EscapedLabel = escapedLabel;
        Number = number;
        ReferenceId = referenceId;
        ReferenceIndex = referenceIndex;
    }

    internal string Label { get; }
    internal string EscapedLabel { get; }
    internal int Number { get; }
    internal string ReferenceId { get; }
    internal int ReferenceIndex { get; }
}

internal struct HtmlFootnoteDefinitionInfo {
    internal HtmlFootnoteDefinitionInfo(
        string label,
        string escapedLabel,
        int number,
        IReadOnlyList<string> referenceIds) {
        Label = label;
        EscapedLabel = escapedLabel;
        Number = number;
        ReferenceIds = referenceIds;
    }

    internal string Label { get; }
    internal string EscapedLabel { get; }
    internal int Number { get; }
    internal IReadOnlyList<string> ReferenceIds { get; }
}
