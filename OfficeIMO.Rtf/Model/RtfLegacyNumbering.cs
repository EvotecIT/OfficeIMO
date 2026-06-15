namespace OfficeIMO.Rtf;

/// <summary>
/// Word 6/95 legacy paragraph numbering metadata represented by the <c>{\*\pn ...}</c> destination and related <c>\pn*</c> controls.
/// </summary>
public sealed class RtfLegacyNumbering {
    /// <summary>Whether legacy paragraph numbering is explicitly enabled by <c>\pn</c>.</summary>
    public bool Enabled { get; set; }

    /// <summary>Legacy numbering level kind.</summary>
    public RtfLegacyNumberingLevelKind LevelKind { get; set; } = RtfLegacyNumberingLevelKind.None;

    /// <summary>Explicit one-based paragraph level from <c>\pnlvl</c>, when present.</summary>
    public int? Level { get; set; }

    /// <summary>Legacy numbering text style.</summary>
    public RtfLegacyNumberingStyle NumberStyle { get; set; } = RtfLegacyNumberingStyle.None;

    /// <summary>Number each table cell only once, represented by <c>\pnnumonce</c>.</summary>
    public bool? NumberEachCellOnce { get; set; }

    /// <summary>Number across rows instead of down columns, represented by <c>\pnacross</c>.</summary>
    public bool? NumberAcrossRows { get; set; }

    /// <summary>Paragraph uses a hanging indent, represented by <c>\pnhang</c>.</summary>
    public bool? HangingIndent { get; set; }

    /// <summary>Restart numbering after section breaks, represented by <c>\pnrestart</c>.</summary>
    public bool? RestartAfterSection { get; set; }

    /// <summary>Include previous level numbers, represented by <c>\pnprev</c>.</summary>
    public bool? IncludePreviousLevels { get; set; }

    /// <summary>Minimum distance from margin to body text in twips, represented by <c>\pnindent</c>.</summary>
    public int? IndentTwips { get; set; }

    /// <summary>Distance from number text to body text in twips, represented by <c>\pnsp</c>.</summary>
    public int? SpaceTwips { get; set; }

    /// <summary>Starting number represented by <c>\pnstart</c>.</summary>
    public int? StartAt { get; set; }

    /// <summary>Numbering alignment represented by <c>\pnql</c>, <c>\pnqc</c>, or <c>\pnqr</c>.</summary>
    public RtfLegacyNumberingAlignment? Alignment { get; set; }

    /// <summary>Font id for the generated number text, represented by <c>\pnf</c>.</summary>
    public int? FontId { get; set; }

    /// <summary>Font size for generated number text in half-points, represented by <c>\pnfs</c>.</summary>
    public int? FontSizeHalfPoints { get; set; }

    /// <summary>Foreground color table index for generated number text, represented by <c>\pncf</c>.</summary>
    public int? ForegroundColorIndex { get; set; }

    /// <summary>Whether generated number text is bold, represented by <c>\pnb</c>.</summary>
    public bool? Bold { get; set; }

    /// <summary>Whether generated number text is italic, represented by <c>\pni</c>.</summary>
    public bool? Italic { get; set; }

    /// <summary>Whether generated number text uses all caps, represented by <c>\pncaps</c>.</summary>
    public bool? AllCaps { get; set; }

    /// <summary>Whether generated number text uses small caps, represented by <c>\pnscaps</c>.</summary>
    public bool? SmallCaps { get; set; }

    /// <summary>Underline style for generated number text.</summary>
    public RtfUnderlineStyle? UnderlineStyle { get; set; }

    /// <summary>Whether generated number text is struck through, represented by <c>\pnstrike</c>.</summary>
    public bool? Strike { get; set; }

    /// <summary>Text preceding the generated number, represented by the <c>\pntxtb</c> destination.</summary>
    public string? TextBefore { get; set; }

    /// <summary>Text following the generated number, represented by the <c>\pntxta</c> destination.</summary>
    public string? TextAfter { get; set; }

    /// <summary>Whether any legacy numbering metadata is present.</summary>
    public bool HasAnyValue =>
        Enabled ||
        LevelKind != RtfLegacyNumberingLevelKind.None ||
        Level.HasValue ||
        NumberStyle != RtfLegacyNumberingStyle.None ||
        NumberEachCellOnce.HasValue ||
        NumberAcrossRows.HasValue ||
        HangingIndent.HasValue ||
        RestartAfterSection.HasValue ||
        IncludePreviousLevels.HasValue ||
        IndentTwips.HasValue ||
        SpaceTwips.HasValue ||
        StartAt.HasValue ||
        Alignment.HasValue ||
        FontId.HasValue ||
        FontSizeHalfPoints.HasValue ||
        ForegroundColorIndex.HasValue ||
        Bold.HasValue ||
        Italic.HasValue ||
        AllCaps.HasValue ||
        SmallCaps.HasValue ||
        UnderlineStyle.HasValue ||
        Strike.HasValue ||
        TextBefore != null ||
        TextAfter != null;

    /// <summary>Copies all legacy numbering metadata from another instance.</summary>
    public void CopyFrom(RtfLegacyNumbering source) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        Enabled = source.Enabled;
        LevelKind = source.LevelKind;
        Level = source.Level;
        NumberStyle = source.NumberStyle;
        NumberEachCellOnce = source.NumberEachCellOnce;
        NumberAcrossRows = source.NumberAcrossRows;
        HangingIndent = source.HangingIndent;
        RestartAfterSection = source.RestartAfterSection;
        IncludePreviousLevels = source.IncludePreviousLevels;
        IndentTwips = source.IndentTwips;
        SpaceTwips = source.SpaceTwips;
        StartAt = source.StartAt;
        Alignment = source.Alignment;
        FontId = source.FontId;
        FontSizeHalfPoints = source.FontSizeHalfPoints;
        ForegroundColorIndex = source.ForegroundColorIndex;
        Bold = source.Bold;
        Italic = source.Italic;
        AllCaps = source.AllCaps;
        SmallCaps = source.SmallCaps;
        UnderlineStyle = source.UnderlineStyle;
        Strike = source.Strike;
        TextBefore = source.TextBefore;
        TextAfter = source.TextAfter;
    }

    /// <summary>Clears all legacy numbering metadata.</summary>
    public void Clear() {
        Enabled = false;
        LevelKind = RtfLegacyNumberingLevelKind.None;
        Level = null;
        NumberStyle = RtfLegacyNumberingStyle.None;
        NumberEachCellOnce = null;
        NumberAcrossRows = null;
        HangingIndent = null;
        RestartAfterSection = null;
        IncludePreviousLevels = null;
        IndentTwips = null;
        SpaceTwips = null;
        StartAt = null;
        Alignment = null;
        FontId = null;
        FontSizeHalfPoints = null;
        ForegroundColorIndex = null;
        Bold = null;
        Italic = null;
        AllCaps = null;
        SmallCaps = null;
        UnderlineStyle = null;
        Strike = null;
        TextBefore = null;
        TextAfter = null;
    }
}
