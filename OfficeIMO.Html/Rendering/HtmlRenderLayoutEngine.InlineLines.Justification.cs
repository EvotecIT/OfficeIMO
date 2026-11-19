using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Distributes each justified line's free space over its justification opportunities and returns
    /// the space added per line. Word separators are always opportunities. text-justify: inter-character
    /// adds the gaps between typographic characters, and auto adds the gaps next to CJK characters,
    /// which are set without word separators. Cursive and right-to-left text keeps its characters
    /// together. The last line of the block and lines ending in a forced break follow text-align-last;
    /// with its auto value they align to start.
    /// </summary>
    private double[] ApplyInlineLineJustification(IReadOnlyList<InlineLine> lines, double width,
        HtmlRenderBoxStyle paragraphStyle, out JustificationLimit limits) {
        var added = new double[lines.Count];
        limits = JustificationLimit.None;
        // Right-to-left paragraphs whose lines cannot be bidi-resolved, because they hold
        // atomic boxes or markers, place segments right to left in logical order. Splitting a
        // left-to-right phrase at its separators would reverse its words, so they stay at start.
        bool unresolvedRightToLeft = string.Equals(paragraphStyle.Direction, "rtl", StringComparison.Ordinal)
            && lines.SelectMany(static line => line.Segments).Any(static segment =>
                segment.Run.AtomicBlock != null || segment.Run.IsEmptyInlineBox
                || segment.Run.RunningStringElement != null || segment.Run.PositionedMarkerElement != null);
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            InlineLine line = lines[lineIndex];
            ClearInlineLineJustification(line);
            if (!JustifiesInlineLine(lines, lineIndex, paragraphStyle)) continue;
            if (unresolvedRightToLeft) {
                // Only text that asks for justification is left unjustified; text-justify: none is not.
                if (line.Segments.Any(static segment => AllowsJustification(segment))) limits |= JustificationLimit.RightToLeftWithInlineBoxes;
                continue;
            }
            double extra = line.ResolveAvailableWidth(width) - line.Width;
            if (extra <= 0.01D) continue;

            List<JustificationUnit> units = CollectJustificationUnits(line);
            var opportunities = new List<JustificationUnit>();
            for (int index = 0; index < units.Count - 1; index++) {
                if (IsJustificationOpportunity(units[index], units[index + 1],
                    ResolveGapTextJustifyMode(line, units[index], units[index + 1], paragraphStyle))) opportunities.Add(units[index]);
            }
            if (opportunities.Count == 0) continue;

            double perOpportunity = extra / opportunities.Count;
            foreach (JustificationUnit unit in opportunities) {
                InlineSegment segment = line.Segments[unit.SegmentIndex];
                segment.ClusterJustification ??= new double[unit.ElementCount];
                segment.ClusterJustification[unit.ElementIndex] += perOpportunity;
                segment.JustificationAdvance += perOpportunity;
            }
            added[lineIndex] = extra;
        }
        return added;
    }

    [Flags]
    private enum JustificationLimit {
        None = 0,
        RightToLeftWithInlineBoxes = 1
    }

    // text-justify applies to text: a gap inside one element's text uses its value, and a gap between
    // two elements' text uses the value of their nearest common ancestor, whether or not it owns text.
    private string ResolveGapTextJustifyMode(InlineLine line, JustificationUnit unit, JustificationUnit next,
        HtmlRenderBoxStyle paragraphStyle) {
        HtmlInlineRun run = line.Segments[unit.SegmentIndex].Run;
        HtmlInlineRun nextRun = line.Segments[next.SegmentIndex].Run;
        if (ReferenceEquals(run, nextRun) || ReferenceEquals(run.OwnerElement, nextRun.OwnerElement)) return run.Style.TextJustifyMode;
        AngleSharp.Dom.IElement? ancestor = NearestCommonAncestor(run.OwnerElement, nextRun.OwnerElement);
        if (ancestor == null) return paragraphStyle.TextJustifyMode;
        if (ReferenceEquals(ancestor, run.OwnerElement)) return run.Style.TextJustifyMode;
        if (ReferenceEquals(ancestor, nextRun.OwnerElement)) return nextRun.Style.TextJustifyMode;
        return _styleResolver.ResolveTextJustifyMode(ancestor) ?? paragraphStyle.TextJustifyMode;
    }

    private static AngleSharp.Dom.IElement? NearestCommonAncestor(AngleSharp.Dom.IElement? left, AngleSharp.Dom.IElement? right) {
        if (left == null || right == null) return null;
        var ancestors = new HashSet<AngleSharp.Dom.IElement>();
        for (AngleSharp.Dom.IElement? current = left; current != null; current = current.ParentElement) ancestors.Add(current);
        for (AngleSharp.Dom.IElement? current = right; current != null; current = current.ParentElement) {
            if (ancestors.Contains(current)) return current;
        }
        return null;
    }

    // text-justify applies to the text it is set on, so text with none is not justified and loses nothing.
    private static bool AllowsJustification(InlineSegment segment) =>
        !segment.Run.Style.TextJustifyNone && IsJustificationContent(segment);

    // Justified lines that are approximated are reported once per paragraph and case.
    private void ReportJustificationLimits(JustificationLimit limits, HtmlRenderBoxStyle paragraphStyle,
        AngleSharp.Dom.IElement? formattingContainer) {
        string source = formattingContainer != null ? HtmlRenderStyleResolver.DescribeSource(formattingContainer) : paragraphStyle.SemanticRole;
        if ((limits & JustificationLimit.RightToLeftWithInlineBoxes) != 0)
            ReportJustificationLimit(formattingContainer, source, "text-align=justify;rtl-inline-boxes",
                "Right-to-left lines with inline boxes were not justified; they were aligned to start.");
    }

    private void ReportJustificationLimit(object? owner, string source, string detail, string message) {
        if (!_reportedJustificationLimits.Add((owner ?? source, detail))) return;
        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.CssPropertyUnsupported, message,
            HtmlDiagnosticSeverity.Warning, source, detail, OfficeConversionLossKind.Approximation);
    }

    private readonly HashSet<(object Owner, string Detail)> _reportedJustificationLimits = new();

    private static void ClearInlineLineJustification(InlineLine line) {
        foreach (InlineSegment segment in line.Segments) {
            segment.JustificationAdvance = 0D;
            segment.ClusterJustification = null;
        }
    }

    private enum JustificationUnitKind {
        Separator,
        Character,
        FixedCharacter,
        Box
    }

    private readonly struct JustificationUnit {
        internal JustificationUnit(int segmentIndex, int elementIndex, int elementCount, JustificationUnitKind kind, string text) {
            SegmentIndex = segmentIndex;
            ElementIndex = elementIndex;
            ElementCount = elementCount;
            Kind = kind;
            Text = text;
        }

        internal int SegmentIndex { get; }
        internal int ElementIndex { get; }
        internal int ElementCount { get; }
        internal JustificationUnitKind Kind { get; }
        internal string Text { get; }
    }

    // The line's typographic units between its first and last content, in order.
    private static List<JustificationUnit> CollectJustificationUnits(InlineLine line) {
        int first = -1, last = -1;
        for (int index = 0; index < line.Segments.Count; index++) {
            if (!IsJustificationContent(line.Segments[index])) continue;
            if (first < 0) first = index;
            last = index;
        }
        var units = new List<JustificationUnit>();
        for (int index = first; index >= 0 && index <= last; index++) {
            InlineSegment segment = line.Segments[index];
            if (!InlineLine.ContributesFlowContent(segment.Run)) continue;
            if (segment.Run.AtomicBlock != null || segment.Text.Length == 0) {
                if (segment.Run.AtomicBlock != null) units.Add(new JustificationUnit(index, 0, 0, JustificationUnitKind.Box, string.Empty));
                continue;
            }
            IReadOnlyList<string> elements = OfficeTextElements.Split(segment.Text);
            // text-justify applies to the text it is set on, so text with none keeps its spacing.
            if (segment.Run.Style.TextJustifyNone) {
                for (int element = 0; element < elements.Count; element++)
                    units.Add(new JustificationUnit(index, element, elements.Count, JustificationUnitKind.FixedCharacter, elements[element]));
                continue;
            }
            if (IsWhitespaceToken(segment.Text)) {
                // Collapsed white space is one separator; preserved spaces each count, while tabs keep their stops.
                if (!segment.Run.Style.PreserveWhitespace) {
                    units.Add(new JustificationUnit(index, 0, elements.Count, JustificationUnitKind.Separator, segment.Text));
                    continue;
                }
                for (int element = 0; element < elements.Count; element++) {
                    units.Add(new JustificationUnit(index, element, elements.Count,
                        elements[element] is " " or " " ? JustificationUnitKind.Separator : JustificationUnitKind.FixedCharacter,
                        elements[element]));
                }
                continue;
            }
            // Shaped, cursive and right-to-left text is not split between characters. A no-break
            // space inside a word does not break the line but still separates words for justification.
            bool fixedCharacters = segment.PreparedShaping != null
                || OfficeTextElements.ContainsRightToLeft(segment.Text)
                || OfficeTextElements.ContainsJoiningScript(segment.Text);
            for (int element = 0; element < elements.Count; element++) {
                units.Add(new JustificationUnit(index, element, elements.Count,
                    fixedCharacters ? JustificationUnitKind.FixedCharacter
                        : elements[element] == "\u00A0" ? JustificationUnitKind.Separator
                        : JustificationUnitKind.Character, elements[element]));
            }
        }
        return units;
    }

    private static bool IsJustificationOpportunity(JustificationUnit unit, JustificationUnit next, string mode) {
        // A word separator belongs to its own text, which already made it fixed when that text has none.
        if (unit.Kind == JustificationUnitKind.Separator) return true;
        // Between two characters, the governing element's none turns the gap off.
        if (mode == "none") return false;
        // A following separator already owns the gap.
        if (next.Kind == JustificationUnitKind.Separator) return false;
        if (mode == "inter-word") return false;
        if (unit.Kind != JustificationUnitKind.Character) return false;
        if (mode == "inter-character") return true;
        return next.Kind == JustificationUnitKind.Character && (IsCjkCharacter(unit.Text) || IsCjkCharacter(next.Text));
    }

    // Han, kana, bopomofo, CJK symbols and fullwidth forms, which CJK text justifies between.
    private static bool IsCjkCharacter(string element) {
        if (element.Length == 0) return false;
        int value = char.IsHighSurrogate(element[0]) && element.Length > 1
            ? char.ConvertToUtf32(element[0], element[1])
            : element[0];
        return value is >= 0x2E80 and <= 0x2FDF
            or >= 0x3000 and <= 0x30FF
            or >= 0x3100 and <= 0x312F
            or >= 0x31A0 and <= 0x31FF
            or >= 0x3400 and <= 0x4DBF
            or >= 0x4E00 and <= 0x9FFF
            or >= 0xF900 and <= 0xFAFF
            or >= 0xFF00 and <= 0xFF60
            or >= 0x20000 and <= 0x3FFFF;
    }

    /// <summary>The alignment of a line that is not justified, or the fallback when it has no opportunity.</summary>
    private static OfficeTextAlignment ResolveInlineLineAlignment(IReadOnlyList<InlineLine> lines, int lineIndex,
        HtmlRenderBoxStyle paragraphStyle) {
        if (IsLastInlineLine(lines, lineIndex) && paragraphStyle.TextAlignLast is OfficeTextAlignment last) {
            return last != OfficeTextAlignment.Justify ? last
                : string.Equals(paragraphStyle.Direction, "rtl", StringComparison.Ordinal)
                    ? OfficeTextAlignment.Right : OfficeTextAlignment.Left;
        }
        // Justified text stores its start alignment for lines that are not justified.
        return paragraphStyle.Alignment;
    }

    private static bool JustifiesInlineLine(IReadOnlyList<InlineLine> lines, int lineIndex, HtmlRenderBoxStyle paragraphStyle) =>
        IsLastInlineLine(lines, lineIndex)
            ? paragraphStyle.TextAlignLast == OfficeTextAlignment.Justify
            : paragraphStyle.JustifyText;

    private static bool IsLastInlineLine(IReadOnlyList<InlineLine> lines, int lineIndex) =>
        lineIndex == lines.Count - 1 || lines[lineIndex].HasForcedBreak;

    private static bool IsJustificationContent(InlineSegment segment) =>
        InlineLine.ContributesFlowContent(segment.Run)
        && (segment.Run.AtomicBlock != null || segment.Text.Length > 0 && !IsWhitespaceToken(segment.Text));

    /// <summary>
    /// Places the space a justified segment received. A single character, such as a word separator,
    /// is painted across its whole advance so decorations stay continuous; longer text is painted
    /// one character at a time with its added space. Paint that letter- or word-spacing already
    /// split keeps its pieces, each moved by the space added before it.
    /// </summary>
    private IReadOnlyList<InlinePaintSegment> ApplyJustificationToPaint(InlineSegment segment,
        IReadOnlyList<InlinePaintSegment> segments) {
        double[]? spacing = segment.ClusterJustification;
        if (spacing == null || segment.JustificationAdvance <= 0D || segments.Count == 0) return segments;
        if (spacing.Length > 1 && TrySplitJustifiedPaint(segment, segments, spacing) is { } spaced) return spaced;
        var widened = new InlinePaintSegment[segments.Count];
        for (int index = 0; index < segments.Count; index++) widened[index] = segments[index];
        InlinePaintSegment last = widened[widened.Length - 1];
        widened[widened.Length - 1] = new InlinePaintSegment(last.Text, last.X, last.Width + segment.JustificationAdvance,
            last.Advance + segment.JustificationAdvance, last.LogicalOrder, last.BidiResolved, last.PreparedShaping);
        return widened;
    }

    // Distributes per-character space over unshaped paint pieces that cover the segment's text in
    // order. Returns null when the pieces were shaped or reordered.
    private InlinePaintSegment[]? TrySplitJustifiedPaint(InlineSegment segment,
        IReadOnlyList<InlinePaintSegment> segments, double[] spacing) {
        var pieceElements = new IReadOnlyList<string>[segments.Count];
        int elementCount = 0;
        for (int index = 0; index < segments.Count; index++) {
            if (segments[index].PreparedShaping != null) return null;
            if (index > 0 && segments[index].X < segments[index - 1].X) return null;
            pieceElements[index] = OfficeTextElements.Split(segments[index].Text);
            elementCount += pieceElements[index].Count;
        }
        if (elementCount != spacing.Length) return null;

        HtmlRenderBoxStyle style = segment.Run.Style;
        OfficeFontInfo font = GetEffectiveTextFont(style);
        var spaced = new List<InlinePaintSegment>(elementCount);
        double shift = 0D;
        int element = 0;
        for (int index = 0; index < segments.Count; index++) {
            InlinePaintSegment piece = segments[index];
            IReadOnlyList<string> elements = pieceElements[index];
            if (elements.Count == 1) {
                double advance = piece.Advance + spacing[element];
                spaced.Add(new InlinePaintSegment(piece.Text, piece.X + shift, Math.Max(0.01D, advance), advance,
                    element, piece.BidiResolved));
                shift += spacing[element++];
                continue;
            }
            // A piece of several characters is painted one character at a time; its letter
            // spacing is zero, or spacing would already have split it.
            double cursor = piece.X + shift;
            for (int part = 0; part < elements.Count; part++) {
                double glyphWidth = MeasureSpacedPaintText(elements[part], style, font);
                double advance = glyphWidth + (IsWordSpacingSeparator(elements[part]) ? style.WordSpacing : 0D) + spacing[element];
                spaced.Add(new InlinePaintSegment(elements[part], cursor, Math.Max(0.01D, advance), advance,
                    element, piece.BidiResolved));
                cursor += advance;
                shift += spacing[element++];
            }
        }
        return spaced.ToArray();
    }
}
