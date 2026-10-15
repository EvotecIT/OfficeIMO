using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>
    /// Visual positions of one line's words under the Unicode Bidirectional Algorithm. Words and
    /// their leading gaps stay in logical content order (reading order, tagging and ActualText);
    /// only their x offsets follow the visual order. Each word's glyphs are already ordered by the
    /// text shaper.
    /// </summary>
    private sealed class BidiLineLayout {
        internal BidiLineLayout(double[] visualX, double[] gapX) {
            VisualX = visualX;
            GapX = gapX;
        }

        /// <summary>Offset of each segment's text from the line start.</summary>
        internal double[] VisualX { get; }

        /// <summary>Offset of each segment's leading gap from the line start.</summary>
        internal double[] GapX { get; }
    }

    // The paragraph level (P2–P3) of a rich paragraph with right-to-left text, else null.
    private static int? ResolveRichParagraphBidiLevel(List<List<RichSeg>> lines) {
        bool rightToLeft = false;
        var text = new StringBuilder();
        foreach (List<RichSeg> line in lines) {
            foreach (RichSeg segment in line) {
                if (segment.InlineElement != null) {
                    text.Append('￼');
                    continue;
                }
                if (segment.LeadingSpace) text.Append(' ');
                text.Append(segment.Text);
                rightToLeft |= segment.TextDirection == OfficeTextDirection.RightToLeft || OfficeUnicodeBidi.ContainsRightToLeft(segment.Text);
            }
        }
        if (!rightToLeft) return null;
        int[] scalars = OfficeUnicodeBidi.ToScalars(text.ToString(), out _);
        return OfficeUnicodeBidi.Resolve(scalars, paragraphLevel: null).ParagraphLevel;
    }

    private static BidiLineLayout? CreateBidiLineLayout(List<RichSeg> segments, int paragraphLevel, double wordSpacing) {
        if (segments.Count == 0) return null;
        foreach (RichSeg segment in segments) {
            // Tab stops and inline objects keep logical placement.
            if (segment.InlineElement != null || segment.LeadingIsTab || segment.LeadingTabLeader != PdfTabLeaderStyle.None) return null;
        }

        var scalars = new List<int>();
        var gapIndex = new int[segments.Count];
        var textStart = new int[segments.Count];
        var textEnd = new int[segments.Count];
        for (int index = 0; index < segments.Count; index++) {
            RichSeg segment = segments[index];
            gapIndex[index] = -1;
            if (segment.LeadingSpace) {
                gapIndex[index] = scalars.Count;
                scalars.Add(' ');
            }
            textStart[index] = scalars.Count;
            scalars.AddRange(OfficeUnicodeBidi.ToScalars(segment.Text, out _));
            if (segment.TextDirection == OfficeTextDirection.RightToLeft && segment.Text.Length > 0 && !OfficeUnicodeBidi.ContainsRightToLeft(segment.Text)) {
                // An explicit right-to-left run without strong characters reads as RTL text.
                for (int position = textStart[index]; position < scalars.Count; position++) scalars[position] = 0x200F;
            }
            textEnd[index] = scalars.Count;
        }

        OfficeUnicodeBidi.Paragraph paragraph = OfficeUnicodeBidi.Resolve(scalars.ToArray(), paragraphLevel);
        byte[] levels = paragraph.GetLineLevels(0, scalars.Count);

        // Items alternate gap/text in logical order; a word takes its lowest character level.
        var itemLevels = new List<int>();
        var itemSegment = new List<int>();
        var itemIsGap = new List<bool>();
        for (int index = 0; index < segments.Count; index++) {
            if (gapIndex[index] >= 0) {
                itemLevels.Add(levels[gapIndex[index]]);
                itemSegment.Add(index);
                itemIsGap.Add(true);
            }
            int level = textEnd[index] > textStart[index] ? int.MaxValue : paragraph.ParagraphLevel;
            for (int position = textStart[index]; position < textEnd[index]; position++) level = Math.Min(level, levels[position]);
            itemLevels.Add(level);
            itemSegment.Add(index);
            itemIsGap.Add(false);
        }

        int[] visualOrder = OfficeUnicodeBidi.ReorderItems(itemLevels);
        bool reordered = false;
        for (int index = 0; index < visualOrder.Length; index++) reordered |= visualOrder[index] != index;
        if (!reordered) return null;

        var visualX = new double[segments.Count];
        var gapX = new double[segments.Count];
        double x = 0D;
        foreach (int item in visualOrder) {
            RichSeg segment = segments[itemSegment[item]];
            if (itemIsGap[item]) {
                gapX[itemSegment[item]] = x;
                x += segment.LeadingAdvance + (segment.LeadingSpaceIsExpandable ? wordSpacing * segment.LeadingSpaceCount : 0D);
            } else {
                visualX[itemSegment[item]] = x;
                x += GetRichSegmentWidth(segment);
            }
        }
        return new BidiLineLayout(visualX, gapX);
    }
}
