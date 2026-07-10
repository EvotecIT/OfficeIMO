namespace OfficeIMO.Html;

internal sealed class HtmlRenderFlowBlock {
    internal HtmlRenderFlowBlock(double width, double height, IEnumerable<HtmlRenderVisual> visuals, bool breakBefore, bool breakAfter, bool avoidBreakInside, string source, IEnumerable<double>? breakOffsets = null) {
        Width = width;
        Height = height;
        Visuals = new List<HtmlRenderVisual>(visuals);
        BreakBefore = breakBefore;
        BreakAfter = breakAfter;
        AvoidBreakInside = avoidBreakInside;
        Source = source;
        var offsets = new SortedSet<double> { 0D, height };
        if (breakOffsets != null) {
            foreach (double offset in breakOffsets) {
                if (offset > 0D && offset < height && !double.IsNaN(offset) && !double.IsInfinity(offset)) offsets.Add(offset);
            }
        }

        BreakOffsets = offsets.ToList().AsReadOnly();
    }

    internal double Width { get; }
    internal double Height { get; }
    internal IReadOnlyList<HtmlRenderVisual> Visuals { get; }
    internal bool BreakBefore { get; }
    internal bool BreakAfter { get; }
    internal bool AvoidBreakInside { get; }
    internal string Source { get; }
    internal IReadOnlyList<double> BreakOffsets { get; }
}

internal sealed class HtmlInlineRun {
    internal HtmlInlineRun(string text, HtmlRenderBoxStyle style, string? linkUri, string source) {
        Text = text;
        Style = style;
        LinkUri = linkUri;
        Source = source;
    }

    internal string Text { get; }
    internal HtmlRenderBoxStyle Style { get; }
    internal string? LinkUri { get; }
    internal string Source { get; }
}

internal sealed class HtmlInlineLayout {
    internal HtmlInlineLayout(IEnumerable<HtmlRenderVisual> visuals, double height, IEnumerable<double>? breakOffsets = null) {
        Visuals = new List<HtmlRenderVisual>(visuals);
        Height = height;
        BreakOffsets = new List<double>(breakOffsets ?? Array.Empty<double>()).AsReadOnly();
    }

    internal IReadOnlyList<HtmlRenderVisual> Visuals { get; }
    internal double Height { get; }
    internal IReadOnlyList<double> BreakOffsets { get; }
}
