using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Word.Legacy;

/// <summary>A recovered paragraph or table in source order.</summary>
public abstract class LegacyWordBlockContent {
    internal LegacyWordBlockContent() { }

    internal static LegacyWordBlockContent From(LegacyWordBlock block) => block switch {
        LegacyWordParagraph paragraph => new LegacyWordParagraphContent(paragraph),
        LegacyWordTable table => new LegacyWordTableContent(table),
        _ => throw new InvalidOperationException("Unknown legacy Word block.")
    };
}

/// <summary>Describes a recovered table. Formatting outside the qualified profile is reported as loss.</summary>
public sealed class LegacyWordTableContent : LegacyWordBlockContent {
    internal LegacyWordTableContent(LegacyWordTable source) {
        ColumnWidthsPoints = source.ColumnWidthsPoints.AsReadOnly();
        Rows = source.Rows.ConvertAll(row => new LegacyWordTableRowContent(row)).AsReadOnly();
    }
    /// <summary>Gets explicitly recovered column widths in points.</summary>
    public IReadOnlyList<double> ColumnWidthsPoints { get; }
    /// <summary>Gets rows in source order.</summary>
    public IReadOnlyList<LegacyWordTableRowContent> Rows { get; }
}

/// <summary>Describes a recovered table row.</summary>
public sealed class LegacyWordTableRowContent {
    internal LegacyWordTableRowContent(LegacyWordTableRow source) {
        IsHeader = source.IsHeader;
        Cells = source.Cells.ConvertAll(cell => new LegacyWordTableCellContent(cell)).AsReadOnly();
    }
    /// <summary>Gets whether the source marks this row as a repeated header.</summary>
    public bool IsHeader { get; }
    /// <summary>Gets cells in source order.</summary>
    public IReadOnlyList<LegacyWordTableCellContent> Cells { get; }
}

/// <summary>Describes a recovered table cell and its formatted paragraphs.</summary>
public sealed class LegacyWordTableCellContent {
    internal LegacyWordTableCellContent(LegacyWordTableCell source) {
        ColumnSpan = source.ColumnSpan;
        Paragraphs = source.Paragraphs.ConvertAll(paragraph => new LegacyWordParagraphContent(paragraph)).AsReadOnly();
    }
    /// <summary>Gets the number of source grid columns occupied by this cell.</summary>
    public int ColumnSpan { get; }
    /// <summary>Gets recovered paragraphs.</summary>
    public IReadOnlyList<LegacyWordParagraphContent> Paragraphs { get; }
}

/// <summary>Identifies pages to which a recovered header or footer applies.</summary>
public enum LegacyWordHeaderFooterOccurrence {
    /// <summary>All pages.</summary>
    All,
    /// <summary>Odd pages.</summary>
    Odd,
    /// <summary>Even pages.</summary>
    Even
}

/// <summary>Describes a recovered running header or footer.</summary>
public sealed class LegacyWordHeaderFooterContent {
    internal LegacyWordHeaderFooterContent(LegacyWordHeaderFooter source) {
        IsFooter = source.IsFooter;
        Occurrence = source.Occurrence;
        Paragraphs = source.Paragraphs.ConvertAll(paragraph => new LegacyWordParagraphContent(paragraph)).AsReadOnly();
    }
    /// <summary>Gets whether this is a footer rather than a header.</summary>
    public bool IsFooter { get; }
    /// <summary>Gets applicable pages.</summary>
    public LegacyWordHeaderFooterOccurrence Occurrence { get; }
    /// <summary>Gets recovered formatted paragraphs.</summary>
    public IReadOnlyList<LegacyWordParagraphContent> Paragraphs { get; }
}

/// <summary>A source-ordered section with recovered page geometry and running stories.</summary>
public sealed class LegacyWordSectionContent {
    internal LegacyWordSectionContent(LegacyWordSection source) {
        Blocks = Array.AsReadOnly(source.Blocks.Select(LegacyWordBlockContent.From).ToArray());
        HeadersAndFooters = source.HeadersAndFooters.ConvertAll(story => new LegacyWordHeaderFooterContent(story)).AsReadOnly();
        PageWidthPoints = source.WidthPoints; PageHeightPoints = source.HeightPoints;
        MarginLeftPoints = source.LeftPoints; MarginRightPoints = source.RightPoints;
        MarginTopPoints = source.TopPoints; MarginBottomPoints = source.BottomPoints;
        StartsNewPage = source.StartsNewPage;
    }
    /// <summary>Gets paragraphs and tables in source order.</summary>
    public IReadOnlyList<LegacyWordBlockContent> Blocks { get; }
    /// <summary>Gets the running stories effective in this section.</summary>
    public IReadOnlyList<LegacyWordHeaderFooterContent> HeadersAndFooters { get; }
    /// <summary>Gets an explicitly recovered page width in points.</summary>
    public double? PageWidthPoints { get; }
    /// <summary>Gets an explicitly recovered page height in points.</summary>
    public double? PageHeightPoints { get; }
    /// <summary>Gets an explicitly recovered left margin in points.</summary>
    public double? MarginLeftPoints { get; }
    /// <summary>Gets an explicitly recovered right margin in points.</summary>
    public double? MarginRightPoints { get; }
    /// <summary>Gets an explicitly recovered top margin in points.</summary>
    public double? MarginTopPoints { get; }
    /// <summary>Gets an explicitly recovered bottom margin in points.</summary>
    public double? MarginBottomPoints { get; }
    /// <summary>Gets whether the section begins on a new page.</summary>
    public bool StartsNewPage { get; }
}
