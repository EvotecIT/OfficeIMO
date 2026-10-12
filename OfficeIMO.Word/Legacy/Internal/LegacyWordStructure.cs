using System.Collections.Generic;

namespace OfficeIMO.Word.Legacy;

/// <summary>Source-order block shared by legacy decoders and the normal Word projection.</summary>
internal abstract class LegacyWordBlock { }

internal sealed class LegacyWordSection {
    internal List<LegacyWordBlock> Blocks { get; } = new();
    internal List<LegacyWordHeaderFooter> HeadersAndFooters { get; } = new();
    internal double? WidthPoints, HeightPoints, LeftPoints, RightPoints, TopPoints, BottomPoints;
    internal bool StartsNewPage;

    internal LegacyWordSection CopySettings() {
        var section = new LegacyWordSection {
            WidthPoints = WidthPoints, HeightPoints = HeightPoints,
            LeftPoints = LeftPoints, RightPoints = RightPoints, TopPoints = TopPoints, BottomPoints = BottomPoints,
            StartsNewPage = StartsNewPage
        };
        section.HeadersAndFooters.AddRange(HeadersAndFooters);
        return section;
    }
}

internal sealed class LegacyWordTable : LegacyWordBlock {
    internal List<double> ColumnWidthsPoints { get; } = new();
    internal List<LegacyWordTableRow> Rows { get; } = new();
}

internal sealed class LegacyWordTableRow {
    internal List<LegacyWordTableCell> Cells { get; } = new();
    internal bool IsHeader;
}

internal sealed class LegacyWordTableCell {
    internal List<LegacyWordParagraph> Paragraphs { get; } = new();
    internal int ColumnSpan = 1;
}

internal sealed class LegacyWordHeaderFooter {
    internal int Slot;
    internal bool IsFooter;
    internal LegacyWordHeaderFooterOccurrence Occurrence;
    internal List<LegacyWordParagraph> Paragraphs { get; } = new();
}
