namespace OfficeIMO.Html;

/// <summary>Internal paint-context identity retained across layout and fragmentation.</summary>
internal sealed class HtmlRenderStackingContext {
    internal HtmlRenderStackingContext(int zIndex, int sourceOrder) {
        ZIndex = zIndex;
        SourceOrder = sourceOrder;
    }
    internal int ZIndex { get; }
    internal int SourceOrder { get; }
}
