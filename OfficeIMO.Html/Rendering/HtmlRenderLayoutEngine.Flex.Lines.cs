namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static List<FlexLine> CreateFlexLines(
        IReadOnlyList<FlexItem> items,
        string flexWrap,
        double contentWidth,
        double gap) {
        var lines = new List<FlexLine>();
        if (items.Count == 0) return lines;
        var current = new FlexLine();
        double used = 0D;
        foreach (FlexItem item in items) {
            double required = item.Basis + (current.Items.Count > 0 ? gap : 0D);
            if (flexWrap != "nowrap" && current.Items.Count > 0 && used + required > contentWidth + 0.0001D) {
                lines.Add(current);
                current = new FlexLine();
                used = 0D;
                required = item.Basis;
            }

            current.Items.Add(item);
            used += required;
        }

        lines.Add(current);
        return lines;
    }

    private void ResolveFlexLineOffsets(
        IReadOnlyList<FlexLine> lines,
        HtmlRenderBoxStyle style,
        double crossSize,
        double gap,
        string source) {
        if (lines.Count == 0) return;
        if (lines.Count == 1) {
            lines[0].CrossSize = crossSize;
            lines[0].CrossOffset = 0D;
            return;
        }

        double used = lines.Sum(line => line.CrossSize) + gap * (lines.Count - 1D);
        double remaining = Math.Max(0D, crossSize - used);
        bool reverse = style.FlexWrap == "wrap-reverse";
        if (reverse && crossSize + 0.0001D < used) {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.FlexValueUnsupported, "Reverse-wrapped overflow used the normal cross direction to keep page coordinates non-negative.", HtmlDiagnosticSeverity.Warning, source, "flex-wrap=wrap-reverse; cross-size-overflow");
            reverse = false;
        }
        string alignment = style.AlignContent == "normal" ? "stretch" : style.AlignContent;
        double start = 0D;
        double between = gap;
        switch (alignment) {
            case "stretch":
                if (remaining > 0D) {
                    double addition = remaining / lines.Count;
                    foreach (FlexLine line in lines) line.CrossSize += addition;
                    remaining = 0D;
                }
                break;
            case "flex-start":
                break;
            case "flex-end":
                start = remaining;
                break;
            case "start":
                start = reverse ? remaining : 0D;
                break;
            case "end":
                start = reverse ? 0D : remaining;
                break;
            case "center":
                start = remaining / 2D;
                break;
            case "space-between":
                between += remaining / (lines.Count - 1D);
                break;
            case "space-around":
                double around = remaining / lines.Count;
                start = around / 2D;
                between += around;
                break;
            case "space-evenly":
                double evenly = remaining / (lines.Count + 1D);
                start = evenly;
                between += evenly;
                break;
            default:
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.FlexValueUnsupported, "An unsupported align-content value used stretch.", HtmlDiagnosticSeverity.Warning, source, "align-content=" + style.AlignContent);
                if (remaining > 0D) {
                    double addition = remaining / lines.Count;
                    foreach (FlexLine line in lines) line.CrossSize += addition;
                    remaining = 0D;
                }
                break;
        }

        double cursor = start;
        foreach (FlexLine line in lines) {
            line.CrossOffset = reverse ? crossSize - cursor - line.CrossSize : cursor;
            cursor += line.CrossSize + between;
        }
    }

    private sealed class FlexLine {
        internal List<FlexItem> Items { get; } = new List<FlexItem>();
        internal double CrossSize { get; set; }
        internal double CrossOffset { get; set; }
    }
}
