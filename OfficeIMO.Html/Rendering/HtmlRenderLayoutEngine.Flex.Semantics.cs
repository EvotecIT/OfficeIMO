namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Retains one source owner for a flex item's ordinary and repeated page fragments.</summary>
    private HtmlRenderFlowBlock ApplyFlexItemSemantics(HtmlRenderFlowBlock block, FlexItem item) {
        if (item.Style.SemanticArtifact) return block;
        // A later column may paint on the first page before an earlier column's
        // final paragraphs arrive. Keep complete item content together in /K.
        string key = "html-flex-item:" + GetSemanticNodeId(item.SourceElement).ToString(System.Globalization.CultureInfo.InvariantCulture)
            + ":" + item.SourceIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);

        return WrapSemanticBlock(block, HtmlRenderSemanticGroupRole.Division, item.Source, key, item.SourceIndex);
    }
}
