namespace OfficeIMO.Rtf;

/// <summary>
/// Minimal semantic representation of an RTF stylesheet entry.
/// </summary>
public sealed class RtfStyle {
    private readonly List<RtfTabStop> _tabStops = new List<RtfTabStop>();

    /// <summary>Creates a stylesheet entry.</summary>
    public RtfStyle(int id, string name, RtfStyleKind kind = RtfStyleKind.Paragraph) {
        Id = id;
        Name = string.IsNullOrWhiteSpace(name) ? "Style " + id.ToString(CultureInfo.InvariantCulture) : name;
        Kind = kind;
    }

    /// <summary>RTF style id.</summary>
    public int Id { get; }

    /// <summary>Human-readable style name.</summary>
    public string Name { get; set; }

    /// <summary>Style kind.</summary>
    public RtfStyleKind Kind { get; set; }

    /// <summary>Optional base style id.</summary>
    public int? BasedOnStyleId { get; set; }

    /// <summary>Optional next paragraph style id.</summary>
    public int? NextStyleId { get; set; }

    /// <summary>Optional linked style id.</summary>
    public int? LinkedStyleId { get; set; }

    /// <summary>Optional shortcut key metadata from the stylesheet <c>{\*\keycode ...}</c> group.</summary>
    public RtfStyleKeyCode? KeyCode { get; set; }

    /// <summary>Whether the style is marked as additive.</summary>
    public bool Additive { get; set; }

    /// <summary>Whether this style automatically updates from document content.</summary>
    public bool AutoUpdate { get; set; }

    /// <summary>Whether this style is hidden from style lists.</summary>
    public bool Hidden { get; set; }

    /// <summary>Whether this style is locked against modification.</summary>
    public bool Locked { get; set; }

    /// <summary>Whether this style is marked as a personal e-mail style.</summary>
    public bool Personal { get; set; }

    /// <summary>Whether this style is marked as an e-mail compose style.</summary>
    public bool Compose { get; set; }

    /// <summary>Whether this style is marked as an e-mail reply style.</summary>
    public bool Reply { get; set; }

    /// <summary>Whether this style is semi-hidden.</summary>
    public bool SemiHidden { get; set; }

    /// <summary>Whether this style is unhidden after use.</summary>
    public bool UnhideWhenUsed { get; set; }

    /// <summary>Whether this style appears in quick style galleries.</summary>
    public bool QuickFormat { get; set; }

    /// <summary>Optional style priority.</summary>
    public int? Priority { get; set; }

    /// <summary>Optional style revision save id.</summary>
    public int? RevisionSaveId { get; set; }

    /// <summary>Optional direct bold setting carried by the stylesheet entry.</summary>
    public bool? Bold { get; set; }

    /// <summary>Optional direct italic setting carried by the stylesheet entry.</summary>
    public bool? Italic { get; set; }

    /// <summary>Optional direct underline setting carried by the stylesheet entry.</summary>
    public RtfUnderlineStyle? UnderlineStyle { get; set; }

    /// <summary>Optional direct font size in points carried by the stylesheet entry.</summary>
    public double? FontSize { get; set; }

    /// <summary>Optional direct font id carried by the stylesheet entry.</summary>
    public int? FontId { get; set; }

    /// <summary>Optional direct foreground color table index carried by the stylesheet entry.</summary>
    public int? ForegroundColorIndex { get; set; }

    /// <summary>Optional direct highlight color table index carried by the stylesheet entry.</summary>
    public int? HighlightColorIndex { get; set; }

    /// <summary>Paragraph tab stops carried by the stylesheet entry.</summary>
    public IReadOnlyList<RtfTabStop> TabStops => _tabStops.AsReadOnly();

    /// <summary>Optional direct paragraph alignment carried by the stylesheet entry.</summary>
    public RtfTextAlignment? ParagraphAlignment { get; set; }

    /// <summary>Optional direct paragraph text direction carried by the stylesheet entry.</summary>
    public RtfTextDirection? ParagraphDirection { get; set; }

    /// <summary>Optional left paragraph indentation in twips carried by the stylesheet entry.</summary>
    public int? LeftIndentTwips { get; set; }

    /// <summary>Optional right paragraph indentation in twips carried by the stylesheet entry.</summary>
    public int? RightIndentTwips { get; set; }

    /// <summary>Optional first-line paragraph indentation in twips carried by the stylesheet entry.</summary>
    public int? FirstLineIndentTwips { get; set; }

    /// <summary>Optional space before paragraph in twips carried by the stylesheet entry.</summary>
    public int? SpaceBeforeTwips { get; set; }

    /// <summary>Optional space after paragraph in twips carried by the stylesheet entry.</summary>
    public int? SpaceAfterTwips { get; set; }

    /// <summary>Optional automatic space-before setting carried by the stylesheet entry.</summary>
    public bool? SpaceBeforeAuto { get; set; }

    /// <summary>Optional automatic space-after setting carried by the stylesheet entry.</summary>
    public bool? SpaceAfterAuto { get; set; }

    /// <summary>Optional raw RTF line spacing value carried by the stylesheet entry.</summary>
    public int? LineSpacingTwips { get; set; }

    /// <summary>Optional line spacing multiplier flag carried by the stylesheet entry.</summary>
    public bool? LineSpacingMultiple { get; set; }

    /// <summary>Optional paragraph background color table index carried by the stylesheet entry.</summary>
    public int? BackgroundColorIndex { get; set; }

    /// <summary>Optional paragraph pattern foreground color table index carried by the stylesheet entry.</summary>
    public int? ShadingForegroundColorIndex { get; set; }

    /// <summary>Optional raw paragraph shading percentage carried by the stylesheet entry.</summary>
    public int? ShadingPatternPercent { get; set; }

    /// <summary>Optional paragraph shading pattern carried by the stylesheet entry.</summary>
    public RtfShadingPattern ShadingPattern { get; set; } = RtfShadingPattern.None;

    /// <summary>Top paragraph border carried by the stylesheet entry.</summary>
    public RtfParagraphBorder TopBorder { get; } = new RtfParagraphBorder();

    /// <summary>Left paragraph border carried by the stylesheet entry.</summary>
    public RtfParagraphBorder LeftBorder { get; } = new RtfParagraphBorder();

    /// <summary>Bottom paragraph border carried by the stylesheet entry.</summary>
    public RtfParagraphBorder BottomBorder { get; } = new RtfParagraphBorder();

    /// <summary>Right paragraph border carried by the stylesheet entry.</summary>
    public RtfParagraphBorder RightBorder { get; } = new RtfParagraphBorder();

    /// <summary>Optional page-break-before setting carried by the stylesheet entry.</summary>
    public bool? PageBreakBefore { get; set; }

    /// <summary>Optional keep-with-next setting carried by the stylesheet entry.</summary>
    public bool? KeepWithNext { get; set; }

    /// <summary>Optional keep-lines-together setting carried by the stylesheet entry.</summary>
    public bool? KeepLinesTogether { get; set; }

    /// <summary>Optional line-number suppression setting carried by the stylesheet entry.</summary>
    public bool? SuppressLineNumbers { get; set; }

    /// <summary>Optional paragraph automatic hyphenation setting carried by the stylesheet entry.</summary>
    public bool? AutoHyphenation { get; set; }

    /// <summary>Optional contextual spacing setting carried by the stylesheet entry.</summary>
    public bool? ContextualSpacing { get; set; }

    /// <summary>Optional automatic right-indent adjustment setting carried by the stylesheet entry.</summary>
    public bool? AdjustRightIndent { get; set; }

    /// <summary>Optional snap-to-line-grid setting carried by the stylesheet entry.</summary>
    public bool? SnapToLineGrid { get; set; }

    /// <summary>Optional widow/orphan control setting carried by the stylesheet entry.</summary>
    public bool? WidowControl { get; set; }

    /// <summary>Optional paragraph outline level carried by the stylesheet entry.</summary>
    public int? OutlineLevel { get; set; }

    /// <summary>Word 6/95 legacy paragraph numbering metadata carried by this stylesheet entry.</summary>
    public RtfLegacyNumbering LegacyNumbering { get; } = new RtfLegacyNumbering();

    /// <summary>Absolute positioning, text wrapping, and drop-cap metadata carried by the stylesheet entry.</summary>
    public RtfParagraphFrame Frame { get; } = new RtfParagraphFrame();

    /// <summary>Table row and cell formatting carried by a table stylesheet entry.</summary>
    public RtfTableRow TableRowFormat { get; } = new RtfTableRow();

    /// <summary>Adds a paragraph tab stop to this stylesheet entry.</summary>
    public RtfTabStop AddTabStop(int positionTwips, RtfTabAlignment alignment = RtfTabAlignment.Left, RtfTabLeader leader = RtfTabLeader.None) {
        var tabStop = new RtfTabStop(positionTwips, alignment, leader);
        _tabStops.Add(tabStop);
        return tabStop;
    }

    /// <summary>Sets paragraph border formatting for one stylesheet side.</summary>
    public RtfStyle SetBorder(RtfParagraphBorderSide side, RtfParagraphBorderStyle style, int? width = null, int? colorIndex = null) {
        RtfParagraphBorder border = GetBorder(side);
        border.Style = style;
        border.Width = width;
        border.ColorIndex = colorIndex;
        return this;
    }

    /// <summary>Configures absolute positioning frame metadata for this stylesheet entry.</summary>
    public RtfStyle SetFrame(Action<RtfParagraphFrame> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        configure(Frame);
        return this;
    }

    /// <summary>Configures Word 6/95 legacy paragraph numbering metadata for this style.</summary>
    public RtfStyle SetLegacyNumbering(Action<RtfLegacyNumbering> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        LegacyNumbering.Enabled = true;
        configure(LegacyNumbering);
        return this;
    }

    internal void ReplaceTabStops(IEnumerable<RtfTabStop> tabStops) {
        _tabStops.Clear();
        foreach (RtfTabStop tabStop in tabStops) {
            _tabStops.Add(new RtfTabStop(tabStop.PositionTwips, tabStop.Alignment, tabStop.Leader));
        }
    }

    private RtfParagraphBorder GetBorder(RtfParagraphBorderSide side) {
        switch (side) {
            case RtfParagraphBorderSide.Top:
                return TopBorder;
            case RtfParagraphBorderSide.Left:
                return LeftBorder;
            case RtfParagraphBorderSide.Bottom:
                return BottomBorder;
            default:
                return RightBorder;
        }
    }
}
