namespace OfficeIMO.Markdown;

public sealed partial class MarkdownNativeDocument {
    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs the supplied source span.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSourceSpan sourceSpan, out MarkdownSourceSlice slice) =>
        ParseResult.TryCreateSourceSlice(sourceSpan, out slice);

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs a native block source field.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeBlockSourceField field, out MarkdownSourceSlice slice) {
        if (field == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(field.SourceSpan, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs a native inline.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeInline inline, out MarkdownSourceSlice slice) {
        if (inline == null || !inline.SourceSpan.HasValue) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(inline.SourceSpan.Value, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs source-backed inline metadata.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeInlineMetadata metadata, out MarkdownSourceSlice slice) {
        if (metadata == null || !metadata.SourceSpan.HasValue) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(metadata.SourceSpan.Value, out slice);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs the supplied source span when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownSourceSpan sourceSpan, out MarkdownSourceSlice slice) =>
        ParseResult.TryCreateOriginalSourceSlice(sourceSpan, out slice);

    /// <summary>
    /// Creates a source slice over the original reader input that backs the supplied source span when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownSourceSpan sourceSpan,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) =>
        ParseResult.TryCreateOriginalSourceSlice(sourceSpan, out slice, out failureReason);

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native block source field when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeBlockSourceField field, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(field, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native block source field when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeBlockSourceField field,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (field == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        return TryCreateOriginalSourceSlice(field.SourceSpan, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native inline when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeInline inline, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(inline, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native inline when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeInline inline,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (inline == null || !inline.SourceSpan.HasValue) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        return TryCreateOriginalSourceSlice(inline.SourceSpan.Value, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs source-backed inline metadata when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeInlineMetadata metadata, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(metadata, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs source-backed inline metadata when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeInlineMetadata metadata,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (metadata == null || !metadata.SourceSpan.HasValue) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        return TryCreateOriginalSourceSlice(metadata.SourceSpan.Value, out slice, out failureReason);
    }
}
