namespace OfficeIMO.Markdown;

/// <summary>Source information a parse keeps beside its objects rather than on them.</summary>
internal static class MarkdownSourceSideTables {
    internal static void Adopt(MarkdownObject target, MarkdownObject captured) {
        MarkdownGenericAttributeSourceSpans.Adopt(target, captured);
        if (target is MarkdownInline targetInline) {
            MarkdownInlineMetadataSourceSpans.AdoptAuxiliaryMetadata(targetInline, (MarkdownInline)captured);
        }
    }
}
