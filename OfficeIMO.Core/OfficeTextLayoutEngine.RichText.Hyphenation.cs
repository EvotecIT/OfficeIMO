using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextLayoutEngine {
    private static bool TryAddHyphenatedRichTextToken(List<OfficeRichTextLine> lines, RichTextLineBuilder builder,
        RichTextToken token, double width, OfficeTextParagraphIndent indent, OfficeRichTextHyphenation hyphenation,
        double scale, CancellationToken cancellationToken, out bool completed) {
        completed = true;
        bool hasBreak = false;
        foreach (int unused in hyphenation.Breaks(token.SourceStart, token.SourceEnd)) { hasBreak = true; break; }
        if (!hasBreak) return false;
        int start = token.SourceStart;
        while (start < token.SourceEnd) {
            cancellationToken.ThrowIfCancellationRequested();
            RichTextToken remainder = SliceHyphenationToken(token, start, token.SourceEnd);
            double available = Math.Max(0, width - builder.OffsetX);
            if (builder.MeasureWidthAfter(remainder) <= available) {
                builder.Add(remainder); builder.SourceTextEnd = token.SourceEnd; return true;
            }
            int selected = -1;
            if (builder.IsEmpty || available - builder.Width > hyphenation.Zone * scale) {
                foreach (int point in hyphenation.Breaks(start, token.SourceEnd)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    RichTextToken prefix = SliceHyphenationToken(token, start, point);
                    if (builder.MeasureHyphenatedWidthAfter(prefix, LastHyphenationRun(prefix)) <= available) selected = point;
                }
            }
            if (selected >= 0) {
                RichTextToken prefix = SliceHyphenationToken(token, start, selected);
                builder.Add(prefix); builder.AddGeneratedHyphen(LastHyphenationRun(prefix));
                builder.SourceTextEnd = selected; start = selected;
            } else if (builder.IsEmpty) {
                // A preferred break may not fit this narrow region. Preserve
                // grapheme fallback, then retry original-word points on the next line.
                int end = start;
                using var elements = BrokenTokenElements(remainder).GetEnumerator();
                bool current = elements.MoveNext();
                while (current) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var element = elements.Current; current = elements.MoveNext();
                    var following = current ? elements.Current : default;
                    double fitting = builder.MeasureFittingWidthAfter(element.Run, element.Text, following.Run, following.Text);
                    if (fitting > available && !builder.IsEmpty) break;
                    builder.AddFitted(element.Run, element.Text, fitting);
                    end += element.Text.Length; builder.SourceTextEnd = end;
                }
                start = end;
                if (start == token.SourceEnd) return true;
            }
            if (!AddRichTextLine(lines, builder)) { completed = false; return true; }
            builder.SetOffset(ResolveLineOffset(indent, firstVisualLine: false));
        }
        return true;
    }

    private static OfficeRichTextRun LastHyphenationRun(RichTextToken token) =>
        token.Parts == null ? token.Run : token.Parts[token.Parts.Count - 1].Run;

    private static RichTextToken SliceHyphenationToken(RichTextToken token, int start, int end) {
        if (token.Parts == null) return RichTextToken.CreateText(token.Run,
            token.Text.Substring(start - token.SourceStart, end - start), false, start, end);
        var parts = new List<RichTextToken>();
        foreach (var part in token.Parts) {
            int left = Math.Max(start, part.SourceStart), right = Math.Min(end, part.SourceEnd);
            if (left < right) parts.Add(RichTextToken.CreateText(part.Run,
                part.Text.Substring(left - part.SourceStart, right - left), false, left, right));
        }
        return parts.Count == 1 ? parts[0] : RichTextToken.CreateWord(parts);
    }
}
