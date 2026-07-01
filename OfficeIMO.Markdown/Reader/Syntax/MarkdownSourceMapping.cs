namespace OfficeIMO.Markdown;

/// <summary>
/// Describes the normalized source slice for a span and, when available, its exact original-input slice.
/// </summary>
public readonly struct MarkdownSourceMapping {
    internal MarkdownSourceMapping(
        MarkdownSourceSpan sourceSpan,
        MarkdownSourceSlice normalizedSourceSlice,
        bool hasOriginalSource,
        MarkdownSourceSlice originalSourceSlice,
        MarkdownOriginalSourceMappingKind originalSourceMappingKind,
        MarkdownOriginalSourceSliceFailureReason originalSourceFailureReason) {
        SourceSpan = sourceSpan;
        NormalizedSourceSlice = normalizedSourceSlice;
        HasOriginalSource = hasOriginalSource;
        OriginalSourceSlice = originalSourceSlice;
        OriginalSourceMappingKind = originalSourceMappingKind;
        OriginalSourceFailureReason = originalSourceFailureReason;
    }

    /// <summary>Source span represented by this mapping.</summary>
    public MarkdownSourceSpan SourceSpan { get; }

    /// <summary>Materialized source slice from the normalized markdown text that backs parser spans.</summary>
    public MarkdownSourceSlice NormalizedSourceSlice { get; }

    /// <summary>Whether <see cref="OriginalSourceSlice"/> contains an exact slice from preserved original reader input.</summary>
    public bool HasOriginalSource { get; }

    /// <summary>Materialized source slice from the original reader input when <see cref="HasOriginalSource"/> is true.</summary>
    public MarkdownSourceSlice OriginalSourceSlice { get; }

    /// <summary>Mapping strategy used for the original source slice, or <see cref="MarkdownOriginalSourceMappingKind.Unavailable"/>.</summary>
    public MarkdownOriginalSourceMappingKind OriginalSourceMappingKind { get; }

    /// <summary>Reason the original source slice is unavailable, or <see cref="MarkdownOriginalSourceSliceFailureReason.None"/>.</summary>
    public MarkdownOriginalSourceSliceFailureReason OriginalSourceFailureReason { get; }

    internal MarkdownSourceMapping WithOriginalFailure(MarkdownOriginalSourceSliceFailureReason failureReason) =>
        new MarkdownSourceMapping(
            SourceSpan,
            NormalizedSourceSlice,
            hasOriginalSource: false,
            originalSourceSlice: default,
            MarkdownOriginalSourceMappingKind.Unavailable,
            failureReason);
}
