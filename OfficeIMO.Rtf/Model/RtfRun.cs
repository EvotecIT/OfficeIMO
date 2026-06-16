namespace OfficeIMO.Rtf;

/// <summary>
/// Text run with character-level formatting.
/// </summary>
public sealed class RtfRun : IRtfInline {
    internal RtfRun(string text) {
        Text = text ?? string.Empty;
    }

    /// <summary>Run text.</summary>
    public string Text { get; set; }

    /// <summary>Whether the run is bold.</summary>
    public bool Bold { get; set; }

    /// <summary>Whether the run is italic.</summary>
    public bool Italic { get; set; }

    /// <summary>Whether the run is underlined. Setting this property uses single underline when enabled.</summary>
    public bool Underline {
        get => UnderlineStyle != RtfUnderlineStyle.None;
        set => UnderlineStyle = value ? RtfUnderlineStyle.Single : RtfUnderlineStyle.None;
    }

    /// <summary>Underline style for the run.</summary>
    public RtfUnderlineStyle UnderlineStyle { get; set; } = RtfUnderlineStyle.None;

    /// <summary>Whether the run is struck through.</summary>
    public bool Strike { get; set; }

    /// <summary>Whether the run uses double strikethrough.</summary>
    public bool DoubleStrike { get; set; }

    /// <summary>Whether the run is hidden text.</summary>
    public bool Hidden { get; set; }

    /// <summary>Whether the run text is outlined.</summary>
    public bool Outline { get; set; }

    /// <summary>Whether the run text has a shadow effect.</summary>
    public bool Shadow { get; set; }

    /// <summary>Whether the run text is embossed.</summary>
    public bool Emboss { get; set; }

    /// <summary>Whether the run text is engraved/imprinted.</summary>
    public bool Imprint { get; set; }

    /// <summary>Optional all-caps or small-caps character effect.</summary>
    public RtfCapsStyle CapsStyle { get; set; } = RtfCapsStyle.None;

    /// <summary>Optional superscript or subscript positioning.</summary>
    public RtfVerticalPosition VerticalPosition { get; set; } = RtfVerticalPosition.Baseline;

    /// <summary>Optional font size in points.</summary>
    public double? FontSize { get; set; }

    /// <summary>Optional font id.</summary>
    public int? FontId { get; set; }

    /// <summary>Optional foreground color index.</summary>
    public int? ForegroundColorIndex { get; set; }

    /// <summary>Optional background highlight color index.</summary>
    public int? HighlightColorIndex { get; set; }

    /// <summary>One-based color table index used by character background shading.</summary>
    public int? CharacterBackgroundColorIndex { get; set; }

    /// <summary>One-based color table index used by character pattern foreground shading.</summary>
    public int? CharacterShadingForegroundColorIndex { get; set; }

    /// <summary>Raw RTF <c>\chshdng</c> value, where 10000 represents 100 percent.</summary>
    public int? CharacterShadingPatternPercent { get; set; }

    /// <summary>Named RTF character shading pattern.</summary>
    public RtfShadingPattern CharacterShadingPattern { get; set; } = RtfShadingPattern.None;

    /// <summary>Character border metadata represented by <c>\chbrdr</c> and following border controls.</summary>
    public RtfCharacterBorder CharacterBorder { get; } = new RtfCharacterBorder();

    /// <summary>Optional underline color table index. A null value uses the document default.</summary>
    public int? UnderlineColorIndex { get; set; }

    /// <summary>Optional character spacing in twips. Positive values expand text and negative values condense it.</summary>
    public int? CharacterSpacingTwips { get; set; }

    /// <summary>Optional character scale percentage. A null value uses the document default of 100 percent.</summary>
    public int? CharacterScalePercent { get; set; }

    /// <summary>Optional kerning threshold in half-points.</summary>
    public int? KerningHalfPoints { get; set; }

    /// <summary>Optional signed baseline offset in half-points. Positive values raise text and negative values lower it.</summary>
    public int? CharacterOffsetHalfPoints { get; set; }

    /// <summary>Optional character stylesheet id.</summary>
    public int? StyleId { get; set; }

    /// <summary>Explicit run text direction represented by <c>\ltrch</c> or <c>\rtlch</c>.</summary>
    public RtfTextDirection? Direction { get; set; }

    /// <summary>Optional run language LCID represented by <c>\lang</c>.</summary>
    public int? LanguageId { get; set; }

    /// <summary>Optional hyperlink target. When set, the writer emits the run as an RTF field.</summary>
    public Uri? Hyperlink { get; set; }

    /// <summary>Optional footnote or annotation emitted immediately after this run.</summary>
    public RtfNote? Note { get; set; }

    /// <summary>Optional tracked revision marker for this run.</summary>
    public RtfRevisionKind RevisionKind { get; set; } = RtfRevisionKind.None;

    /// <summary>Optional zero-based index into the document revision author table.</summary>
    public int? RevisionAuthorIndex { get; set; }

    /// <summary>Optional raw RTF DTTM revision timestamp value from <c>\revdttm</c>.</summary>
    public int? RevisionTimestampValue { get; set; }

    /// <summary>Character revision save identifier represented by <c>\charrsid</c>.</summary>
    public int? CharacterRevisionSaveId { get; set; }

    /// <summary>Insertion revision save identifier represented by <c>\insrsid</c>.</summary>
    public int? InsertionRevisionSaveId { get; set; }

    /// <summary>Deletion revision save identifier represented by <c>\delrsid</c>.</summary>
    public int? DeletionRevisionSaveId { get; set; }

    /// <summary>Enables bold formatting.</summary>
    public RtfRun SetBold(bool value = true) {
        Bold = value;
        return this;
    }

    /// <summary>Enables italic formatting.</summary>
    public RtfRun SetItalic(bool value = true) {
        Italic = value;
        return this;
    }

    /// <summary>Enables underline formatting.</summary>
    public RtfRun SetUnderline(bool value = true) {
        Underline = value;
        return this;
    }

    /// <summary>Sets underline formatting to the requested RTF underline style.</summary>
    public RtfRun SetUnderline(RtfUnderlineStyle style) {
        UnderlineStyle = style;
        return this;
    }

    /// <summary>Sets the underline color table index.</summary>
    public RtfRun SetUnderlineColor(int colorIndex) {
        if (colorIndex < 0) throw new ArgumentOutOfRangeException(nameof(colorIndex), "Color index cannot be negative.");
        UnderlineColorIndex = colorIndex;
        return this;
    }

    /// <summary>Sets character spacing in twips.</summary>
    public RtfRun SetCharacterSpacingTwips(int spacingTwips) {
        CharacterSpacingTwips = spacingTwips;
        return this;
    }

    /// <summary>Sets character scale percentage.</summary>
    public RtfRun SetCharacterScale(int percent) {
        if (percent <= 0) throw new ArgumentOutOfRangeException(nameof(percent), "Character scale must be greater than zero.");
        CharacterScalePercent = percent == 100 ? null : percent;
        return this;
    }

    /// <summary>Sets the kerning threshold in half-points.</summary>
    public RtfRun SetKerningHalfPoints(int halfPoints) {
        if (halfPoints < 0) throw new ArgumentOutOfRangeException(nameof(halfPoints), "Kerning threshold cannot be negative.");
        KerningHalfPoints = halfPoints == 0 ? null : halfPoints;
        return this;
    }

    /// <summary>Sets signed baseline offset in half-points.</summary>
    public RtfRun SetCharacterOffsetHalfPoints(int halfPoints) {
        CharacterOffsetHalfPoints = halfPoints == 0 ? null : halfPoints;
        return this;
    }

    /// <summary>Enables strike-through formatting.</summary>
    public RtfRun SetStrike(bool value = true) {
        Strike = value;
        return this;
    }

    /// <summary>Enables double strike-through formatting.</summary>
    public RtfRun SetDoubleStrike(bool value = true) {
        DoubleStrike = value;
        return this;
    }

    /// <summary>Enables hidden text formatting.</summary>
    public RtfRun SetHidden(bool value = true) {
        Hidden = value;
        return this;
    }

    /// <summary>Enables outline text formatting.</summary>
    public RtfRun SetOutline(bool value = true) {
        Outline = value;
        return this;
    }

    /// <summary>Enables shadow text formatting.</summary>
    public RtfRun SetShadow(bool value = true) {
        Shadow = value;
        return this;
    }

    /// <summary>Enables embossed text formatting.</summary>
    public RtfRun SetEmboss(bool value = true) {
        Emboss = value;
        return this;
    }

    /// <summary>Enables engraved/imprinted text formatting.</summary>
    public RtfRun SetImprint(bool value = true) {
        Imprint = value;
        return this;
    }

    /// <summary>Sets the run capitalization effect.</summary>
    public RtfRun SetCapsStyle(RtfCapsStyle capsStyle) {
        CapsStyle = capsStyle;
        return this;
    }

    /// <summary>Marks the run as superscript text.</summary>
    public RtfRun SetSuperscript() {
        VerticalPosition = RtfVerticalPosition.Superscript;
        return this;
    }

    /// <summary>Marks the run as subscript text.</summary>
    public RtfRun SetSubscript() {
        VerticalPosition = RtfVerticalPosition.Subscript;
        return this;
    }

    /// <summary>Resets the run to normal baseline text.</summary>
    public RtfRun SetBaseline() {
        VerticalPosition = RtfVerticalPosition.Baseline;
        return this;
    }

    /// <summary>Sets the run font size in points.</summary>
    public RtfRun SetFontSize(double points) {
        if (points <= 0) throw new ArgumentOutOfRangeException(nameof(points), "Font size must be greater than zero.");
        FontSize = points;
        return this;
    }

    /// <summary>Sets the foreground color table index.</summary>
    public RtfRun SetForegroundColor(int colorIndex) {
        if (colorIndex < 0) throw new ArgumentOutOfRangeException(nameof(colorIndex), "Color index cannot be negative.");
        ForegroundColorIndex = colorIndex;
        return this;
    }

    /// <summary>Sets the highlight color table index.</summary>
    public RtfRun SetHighlightColor(int colorIndex) {
        if (colorIndex < 0) throw new ArgumentOutOfRangeException(nameof(colorIndex), "Color index cannot be negative.");
        HighlightColorIndex = colorIndex;
        return this;
    }

    /// <summary>Sets character background shading to a one-based color table index.</summary>
    public RtfRun SetCharacterBackgroundColor(int? colorIndex) {
        CharacterBackgroundColorIndex = colorIndex;
        return this;
    }

    /// <summary>Sets character shading color and pattern metadata.</summary>
    public RtfRun SetCharacterShading(int? backgroundColorIndex, int? foregroundColorIndex = null, int? patternPercent = null, RtfShadingPattern pattern = RtfShadingPattern.None) {
        CharacterBackgroundColorIndex = backgroundColorIndex;
        CharacterShadingForegroundColorIndex = foregroundColorIndex;
        CharacterShadingPatternPercent = patternPercent;
        CharacterShadingPattern = pattern;
        return this;
    }

    /// <summary>Sets the character border formatting.</summary>
    public RtfRun SetCharacterBorder(RtfParagraphBorderStyle style, int? width = null, int? colorIndex = null) {
        CharacterBorder.Style = style;
        CharacterBorder.Width = width;
        CharacterBorder.ColorIndex = colorIndex;
        return this;
    }

    /// <summary>Sets the character style id.</summary>
    public RtfRun SetStyle(int styleId) {
        StyleId = styleId;
        return this;
    }

    /// <summary>Sets explicit run text direction.</summary>
    public RtfRun SetDirection(RtfTextDirection? direction) {
        Direction = direction;
        return this;
    }

    /// <summary>Sets the run language LCID.</summary>
    public RtfRun SetLanguage(int languageId) {
        if (languageId < 0) throw new ArgumentOutOfRangeException(nameof(languageId), "Language id cannot be negative.");
        LanguageId = languageId;
        return this;
    }

    /// <summary>Marks the run as a hyperlink.</summary>
    public RtfRun SetHyperlink(Uri uri) {
        Hyperlink = uri ?? throw new ArgumentNullException(nameof(uri));
        return this;
    }

    /// <summary>Attaches a footnote or annotation to this run.</summary>
    public RtfRun SetNote(RtfNote note) {
        Note = note ?? throw new ArgumentNullException(nameof(note));
        return this;
    }

    /// <summary>Marks the run as tracked inserted text.</summary>
    public RtfRun SetInsertedRevision(int? authorIndex = null, int? timestampValue = null) {
        RevisionKind = RtfRevisionKind.Inserted;
        RevisionAuthorIndex = authorIndex;
        RevisionTimestampValue = timestampValue;
        return this;
    }

    /// <summary>Marks the run as tracked deleted text.</summary>
    public RtfRun SetDeletedRevision(int? authorIndex = null, int? timestampValue = null) {
        RevisionKind = RtfRevisionKind.Deleted;
        RevisionAuthorIndex = authorIndex;
        RevisionTimestampValue = timestampValue;
        return this;
    }

    /// <summary>Clears tracked revision metadata from the run.</summary>
    public RtfRun ClearRevision() {
        RevisionKind = RtfRevisionKind.None;
        RevisionAuthorIndex = null;
        RevisionTimestampValue = null;
        return this;
    }

    /// <summary>Sets run-level revision save identifiers.</summary>
    public RtfRun SetRevisionSaveIds(int? character = null, int? insertion = null, int? deletion = null) {
        ValidateRevisionSaveId(character, nameof(character));
        ValidateRevisionSaveId(insertion, nameof(insertion));
        ValidateRevisionSaveId(deletion, nameof(deletion));
        CharacterRevisionSaveId = character;
        InsertionRevisionSaveId = insertion;
        DeletionRevisionSaveId = deletion;
        return this;
    }

    private static void ValidateRevisionSaveId(int? id, string parameterName) {
        if (id.HasValue && id.Value < 0) {
            throw new ArgumentOutOfRangeException(parameterName, "Revision save id cannot be negative.");
        }
    }
}
