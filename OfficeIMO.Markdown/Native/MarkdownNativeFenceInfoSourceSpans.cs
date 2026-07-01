namespace OfficeIMO.Markdown;

internal static class MarkdownNativeFenceInfoSourceSpans {
    internal static string? GetAttributeSourceText(MarkdownCodeFenceInfo? fenceInfo) {
        if (fenceInfo == null ||
            !fenceInfo.HasExplicitAttributes ||
            fenceInfo.GenericAttributes.IsEmpty ||
            string.IsNullOrWhiteSpace(fenceInfo.GenericAttributeSourceText)) {
            return null;
        }

        return fenceInfo.GenericAttributeSourceText;
    }

    internal static MarkdownSourceSpan? GetAttributeSourceSpan(MarkdownCodeFenceInfo? fenceInfo, MarkdownSourceSpan? infoStringSourceSpan) {
        var sourceText = GetAttributeSourceText(fenceInfo);
        if (string.IsNullOrEmpty(sourceText) ||
            fenceInfo == null ||
            !infoStringSourceSpan.HasValue ||
            !infoStringSourceSpan.Value.StartColumn.HasValue ||
            !infoStringSourceSpan.Value.EndColumn.HasValue) {
            return null;
        }

        var infoString = fenceInfo.InfoString ?? string.Empty;
        var relativeStart = infoString.IndexOf(sourceText, StringComparison.Ordinal);
        if (relativeStart < 0) {
            return null;
        }

        var span = infoStringSourceSpan.Value;
        var startColumn = span.StartColumn.Value + relativeStart;
        var endColumn = startColumn + sourceText!.Length - 1;
        int? startOffset = span.StartOffset.HasValue ? span.StartOffset.Value + relativeStart : null;
        int? endOffset = startOffset.HasValue ? startOffset.Value + sourceText.Length - 1 : null;
        return new MarkdownSourceSpan(span.StartLine, startColumn, span.StartLine, endColumn, startOffset, endOffset);
    }
}
