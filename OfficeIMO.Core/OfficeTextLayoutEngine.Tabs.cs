using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextLayoutEngine {
    // Measurements include every styled run in the following field, including delimiters
    // split across run boundaries. Each tab starts a new kerning/measurement context.
    private static double ResolveTabAdvance(OfficeTextTabStops tabs, double current,
        IReadOnlyList<RichTextToken> tokens, int tabIndex,
        Func<string?, double, string?, OfficeFontStyle, double> measure, CancellationToken cancellationToken, out OfficeTextTabStop? selectedStop) {
        selectedStop = null;
        foreach (OfficeTextTabStop stop in tabs.Stops) {
            cancellationToken.ThrowIfCancellationRequested();
            double anchor = tabs.Origin + stop.Position;
            if (anchor <= current + .000001D) continue;
            double alignedWidth = 0;
            if (stop.Alignment != OfficeTextTabAlignment.Left) {
                alignedWidth = MeasureTabField(tokens, tabIndex,
                    stop.Alignment == OfficeTextTabAlignment.Character ? stop.Character : null, measure, cancellationToken);
                if (stop.Alignment == OfficeTextTabAlignment.Center) alignedWidth /= 2;
            }
            double start = anchor - alignedWidth;
            // The next stop is consumed even when right/center/character alignment
            // would overlap preceding text. It advances zero rather than selecting
            // a later stop; subsequent tabs resolve from the resulting text cursor.
            selectedStop = stop;
            return Math.Max(0, start - current);
        }
        double last = tabs.Stops.Count == 0 ? 0 : tabs.Stops[tabs.Stops.Count - 1].Position;
        double coordinate = Math.Max(current - tabs.Origin, last);
        double next = (Math.Floor(coordinate / tabs.DefaultInterval) + 1) * tabs.DefaultInterval + tabs.Origin;
        if (double.IsNaN(next) || double.IsInfinity(next)) return double.PositiveInfinity;
        return Math.Max(0, next - current);
    }

    private static double MeasureTabField(IReadOnlyList<RichTextToken> tokens, int tabIndex, string? character,
        Func<string?, double, string?, OfficeFontStyle, double> measure, CancellationToken cancellationToken) {
        var text = new StringBuilder();
        OfficeRichTextSegment? context = null;
        double width = 0;
        for (int i = tabIndex + 1; i < tokens.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            RichTextToken token = tokens[i];
            if (token.IsTab || token.HardBreak) break;
            int delimiter = character == null ? -1 : token.Text.IndexOf(character, StringComparison.Ordinal);
            int length = delimiter < 0 ? token.Text.Length : delimiter;
            if (length > 0) {
                if (context != null && !RichTextLineBuilder.CanMerge(context, token.Run)) {
                    width += MeasureContext();
                    text.Clear();
                    context = null;
                }
                // Keep the first token's style as the context; merge rules match painting.
                context ??= RichTextLineBuilder.CreateSegment(token.Run, token.Text, 0);
                text.Append(token.Text, 0, length);
            }
            if (delimiter >= 0) break;
        }
        return width + MeasureContext();

        double MeasureContext() {
            cancellationToken.ThrowIfCancellationRequested();
            return context == null ? 0 : Measure(text.ToString(), EffectiveFontSize(context), context.FontFamily, context.FontStyle, measure);
        }
    }
}
