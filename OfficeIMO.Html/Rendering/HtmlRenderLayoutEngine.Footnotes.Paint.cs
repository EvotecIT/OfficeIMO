using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Paints a note fragment in the supplied page or column area, preserving its semantics and bidirectional destinations.</summary>
    private void AddFootnoteChunkVisuals(
        ICollection<HtmlRenderVisual> target, HtmlFootnoteEntry entry, HtmlFootnoteChunk chunk,
        double x, double cursorY, double width) {
        string marker = entry.MarkerContent
            + (chunk.Start > 0.0001D ? "\u00a0(cont.)" : string.Empty);
        double markerGutter = entry.MarkerGutter;
        IReadOnlyList<HtmlRenderVisual> body = SliceBlockVisuals(entry.Block, chunk.Start, chunk.End);
        var children = new List<HtmlRenderVisual>();
        if (chunk.Start <= 0.0001D) {
            children.Add(new HtmlRenderNamedDestination(
                FootnoteNoteDestination(entry.Number),
                x,
                cursorY,
                _paintOrder++,
                HtmlRenderStyleResolver.DescribeSource(entry.Element) + ":footnote-destination"));
        }
        if (!entry.SuppressMarker) {
            if (entry.MarkerBlock != null) {
                foreach (HtmlRenderVisual visual in entry.MarkerBlock.Visuals) {
                    children.Add(visual.Translate(x, cursorY, _paintOrder++));
                }
                if (chunk.Start > 0.0001D) {
                    _fontUsage?.Observe("\u00a0(cont.)", entry.MarkerStyle.Font.FamilyName, entry.MarkerStyle.FontDescriptor);
                    children.Add(new HtmlRenderText(
                        "\u00a0(cont.)",
                        x + entry.MarkerBlock.Width,
                        cursorY,
                        Math.Max(0.01D, markerGutter - entry.MarkerBlock.Width),
                        Math.Max(0.01D, entry.MarkerStyle.LineHeight),
                        entry.MarkerStyle.Font,
                        entry.MarkerStyle.Color,
                        OfficeTextAlignment.Left,
                        entry.MarkerStyle.LineHeight,
                        _paintOrder++,
                        "#" + FootnoteCallDestination(entry.Number),
                        HtmlRenderStyleResolver.DescribeSource(entry.Element) + ":footnote-marker",
                        semanticRole: null,
                        layoutY: null,
                        semanticNodeId: null,
                        textAdvanceWidth: null,
                        featureSettings: entry.MarkerStyle.TextFeatureSettings,
                        fontPalette: entry.MarkerStyle.FontPalette,
                        fontDescriptor: entry.MarkerStyle.FontDescriptor));
                }
            } else {
                children.Add(new HtmlRenderText(
                    marker,
                    x,
                    cursorY,
                    Math.Max(0.01D, markerGutter - 2D),
                    Math.Max(0.01D, entry.MarkerStyle.LineHeight),
                    entry.MarkerStyle.Font,
                    entry.MarkerStyle.Color,
                    OfficeTextAlignment.Left,
                    entry.MarkerStyle.LineHeight,
                    _paintOrder++,
                    "#" + FootnoteCallDestination(entry.Number),
                    HtmlRenderStyleResolver.DescribeSource(entry.Element) + ":footnote-marker",
                    semanticRole: null,
                    layoutY: null,
                    semanticNodeId: null,
                    textAdvanceWidth: null,
                    featureSettings: entry.MarkerStyle.TextFeatureSettings,
                    fontPalette: entry.MarkerStyle.FontPalette,
                    fontDescriptor: entry.MarkerStyle.FontDescriptor));
            }
        }
        foreach (HtmlRenderVisual visual in body) {
            children.Add(visual.Translate(x + markerGutter, cursorY, _paintOrder++));
        }
        target.Add(new HtmlRenderSemanticGroup(
            HtmlRenderSemanticGroupRole.Footnote,
            x,
            cursorY,
            width,
            chunk.ReservedHeight,
            children,
            _paintOrder++,
            HtmlRenderStyleResolver.DescribeSource(entry.Element) + ":footnote",
            structureElementKey: "html-footnote:" + entry.Number.ToString(System.Globalization.CultureInfo.InvariantCulture)));
    }
}
