namespace OfficeIMO.Rtf;

/// <summary>
/// Document-level RTF settings that are not tied to page setup or body content.
/// </summary>
public sealed class RtfDocumentSettings {
    /// <summary>Document character set declaration represented by <c>\ansi</c>, <c>\mac</c>, <c>\pc</c>, or <c>\pca</c>.</summary>
    public RtfDocumentCharacterSet? CharacterSet { get; set; }

    /// <summary>ANSI code page represented by <c>\ansicpg</c>.</summary>
    public int? AnsiCodePage { get; set; }

    /// <summary>Default alternate character count after Unicode escapes, represented by <c>\uc</c>.</summary>
    public int? UnicodeSkipCount { get; set; }

    /// <summary>Default font id represented by <c>\deff</c>.</summary>
    public int? DefaultFontId { get; set; }

    /// <summary>Default tab width in twips, represented by <c>\deftab</c>.</summary>
    public int? DefaultTabWidthTwips { get; set; }

    /// <summary>Default document language LCID represented by <c>\deflang</c>.</summary>
    public int? DefaultLanguageId { get; set; }

    /// <summary>Default Far East language LCID represented by <c>\deflangfe</c>.</summary>
    public int? DefaultFarEastLanguageId { get; set; }

    /// <summary>Default alternate language LCID represented by <c>\adeflang</c>.</summary>
    public int? DefaultAlternateLanguageId { get; set; }

    /// <summary>View kind value represented by <c>\viewkind</c>.</summary>
    public int? ViewKind { get; set; }

    /// <summary>View scale percentage represented by <c>\viewscale</c>.</summary>
    public int? ViewScale { get; set; }

    /// <summary>Zoom kind value represented by <c>\viewzk</c>.</summary>
    public int? ZoomKind { get; set; }

    /// <summary>View backspace behavior represented by <c>\viewbksp</c>.</summary>
    public int? ViewBackspaceBehavior { get; set; }

    /// <summary>Whether widow/orphan control is explicitly enabled or disabled.</summary>
    public bool? WidowOrphanControl { get; set; }

    /// <summary>Whether automatic hyphenation is explicitly enabled or disabled.</summary>
    public bool? AutoHyphenation { get; set; }

    /// <summary>Whether words in all caps may be hyphenated, represented by <c>\hyphcaps</c>.</summary>
    public bool? HyphenateCaps { get; set; }

    /// <summary>Maximum number of consecutive hyphenated lines represented by <c>\hyphconsec</c>.</summary>
    public int? ConsecutiveHyphenLimit { get; set; }

    /// <summary>Hyphenation hot zone in twips represented by <c>\hyphhotz</c>.</summary>
    public int? HyphenationZoneTwips { get; set; }

    /// <summary>Whether facing pages are explicitly enabled or disabled.</summary>
    public bool? FacingPages { get; set; }

    /// <summary>Whether mirrored margins are explicitly enabled or disabled.</summary>
    public bool? MirrorMargins { get; set; }

    /// <summary>Whether form protection is explicitly enabled or disabled.</summary>
    public bool? FormProtection { get; set; }

    /// <summary>Whether revision protection is explicitly enabled or disabled.</summary>
    public bool? RevisionProtection { get; set; }

    /// <summary>Whether annotation protection is explicitly enabled or disabled.</summary>
    public bool? AnnotationProtection { get; set; }

    /// <summary>Whether read-only protection is explicitly enabled or disabled.</summary>
    public bool? ReadOnlyProtection { get; set; }

    /// <summary>Whether revision marking is enabled, represented by <c>\revisions</c>.</summary>
    public bool? TrackRevisions { get; set; }

    /// <summary>Revision text display style represented by <c>\revprop</c>.</summary>
    public int? RevisionDisplayStyle { get; set; }

    /// <summary>Revision bar placement represented by <c>\revbar</c>.</summary>
    public int? RevisionBarPlacement { get; set; }

    /// <summary>Drawing grid horizontal spacing represented by <c>\dghspace</c>.</summary>
    public int? DrawingGridHorizontalSpacingTwips { get; set; }

    /// <summary>Drawing grid vertical spacing represented by <c>\dgvspace</c>.</summary>
    public int? DrawingGridVerticalSpacingTwips { get; set; }

    /// <summary>Drawing grid horizontal origin represented by <c>\dghorigin</c>.</summary>
    public int? DrawingGridHorizontalOriginTwips { get; set; }

    /// <summary>Drawing grid vertical origin represented by <c>\dgvorigin</c>.</summary>
    public int? DrawingGridVerticalOriginTwips { get; set; }

    /// <summary>Drawing grid horizontal display frequency represented by <c>\dghshow</c>.</summary>
    public int? DrawingGridHorizontalShow { get; set; }

    /// <summary>Drawing grid vertical display frequency represented by <c>\dgvshow</c>.</summary>
    public int? DrawingGridVerticalShow { get; set; }

    /// <summary>Whether drawing objects snap to the document grid, represented by <c>\dgsnap</c>.</summary>
    public bool? SnapToDrawingGrid { get; set; }

    /// <summary>Whether drawing grid origin uses page margins, represented by <c>\dgmargin</c>.</summary>
    public bool? DrawingGridUsesMargins { get; set; }

    /// <summary>Document default text direction represented by <c>\ltrdoc</c> or <c>\rtldoc</c>.</summary>
    public RtfTextDirection? Direction { get; set; }

    /// <summary>Sets the default tab width in twips.</summary>
    public RtfDocumentSettings SetDefaultTabWidth(int twips) {
        if (twips < 0) throw new ArgumentOutOfRangeException(nameof(twips), "Default tab width cannot be negative.");
        DefaultTabWidthTwips = twips;
        return this;
    }

    /// <summary>Sets the document character set declaration and optional ANSI code page.</summary>
    public RtfDocumentSettings SetCharacterSet(RtfDocumentCharacterSet characterSet, int? ansiCodePage = null) {
        ValidateNonNegative(ansiCodePage, nameof(ansiCodePage));
        CharacterSet = characterSet;
        AnsiCodePage = ansiCodePage;
        return this;
    }

    /// <summary>Sets the default font id from the document font table.</summary>
    public RtfDocumentSettings SetDefaultFont(int fontId) {
        if (fontId < 0) throw new ArgumentOutOfRangeException(nameof(fontId), "Default font id cannot be negative.");
        DefaultFontId = fontId;
        return this;
    }

    /// <summary>Sets the document default alternate character count after Unicode escapes.</summary>
    public RtfDocumentSettings SetUnicodeSkipCount(int count) {
        if (count < 0) throw new ArgumentOutOfRangeException(nameof(count), "Unicode skip count cannot be negative.");
        UnicodeSkipCount = count;
        return this;
    }

    /// <summary>Sets the default document language LCID.</summary>
    public RtfDocumentSettings SetDefaultLanguage(int languageId) {
        if (languageId < 0) throw new ArgumentOutOfRangeException(nameof(languageId), "Language id cannot be negative.");
        DefaultLanguageId = languageId;
        return this;
    }

    /// <summary>Sets the default Far East document language LCID.</summary>
    public RtfDocumentSettings SetDefaultFarEastLanguage(int languageId) {
        if (languageId < 0) throw new ArgumentOutOfRangeException(nameof(languageId), "Language id cannot be negative.");
        DefaultFarEastLanguageId = languageId;
        return this;
    }

    /// <summary>Sets the default alternate document language LCID.</summary>
    public RtfDocumentSettings SetDefaultAlternateLanguage(int languageId) {
        if (languageId < 0) throw new ArgumentOutOfRangeException(nameof(languageId), "Language id cannot be negative.");
        DefaultAlternateLanguageId = languageId;
        return this;
    }

    /// <summary>Sets the document view controls.</summary>
    public RtfDocumentSettings SetView(int? kind = null, int? scale = null, int? zoomKind = null, int? backspaceBehavior = null) {
        ValidateNonNegative(kind, nameof(kind));
        ValidateNonNegative(scale, nameof(scale));
        ValidateNonNegative(zoomKind, nameof(zoomKind));
        ValidateNonNegative(backspaceBehavior, nameof(backspaceBehavior));
        ViewKind = kind;
        ViewScale = scale;
        ZoomKind = zoomKind;
        ViewBackspaceBehavior = backspaceBehavior;
        return this;
    }

    /// <summary>Sets document-level automatic hyphenation controls.</summary>
    public RtfDocumentSettings SetHyphenation(bool? automatic = null, bool? caps = null, int? consecutiveLimit = null, int? zoneTwips = null) {
        ValidateNonNegative(consecutiveLimit, nameof(consecutiveLimit));
        ValidateNonNegative(zoneTwips, nameof(zoneTwips));
        AutoHyphenation = automatic;
        HyphenateCaps = caps;
        ConsecutiveHyphenLimit = consecutiveLimit;
        HyphenationZoneTwips = zoneTwips;
        return this;
    }

    /// <summary>Sets protection-related flags.</summary>
    public RtfDocumentSettings SetProtection(bool? forms = null, bool? revisions = null, bool? annotations = null, bool? readOnly = null) {
        FormProtection = forms;
        RevisionProtection = revisions;
        AnnotationProtection = annotations;
        ReadOnlyProtection = readOnly;
        return this;
    }

    /// <summary>Sets revision-marking display controls.</summary>
    public RtfDocumentSettings SetRevisionTracking(bool? enabled = null, int? displayStyle = null, int? barPlacement = null) {
        ValidateNonNegative(displayStyle, nameof(displayStyle));
        ValidateNonNegative(barPlacement, nameof(barPlacement));
        TrackRevisions = enabled;
        RevisionDisplayStyle = displayStyle;
        RevisionBarPlacement = barPlacement;
        return this;
    }

    /// <summary>Sets document-level drawing grid controls.</summary>
    public RtfDocumentSettings SetDrawingGrid(
        int? horizontalSpacingTwips = null,
        int? verticalSpacingTwips = null,
        int? horizontalOriginTwips = null,
        int? verticalOriginTwips = null,
        int? horizontalShow = null,
        int? verticalShow = null,
        bool? snapToGrid = null,
        bool? useMargins = null) {
        ValidateNonNegative(horizontalSpacingTwips, nameof(horizontalSpacingTwips));
        ValidateNonNegative(verticalSpacingTwips, nameof(verticalSpacingTwips));
        ValidateNonNegative(horizontalOriginTwips, nameof(horizontalOriginTwips));
        ValidateNonNegative(verticalOriginTwips, nameof(verticalOriginTwips));
        ValidateNonNegative(horizontalShow, nameof(horizontalShow));
        ValidateNonNegative(verticalShow, nameof(verticalShow));
        DrawingGridHorizontalSpacingTwips = horizontalSpacingTwips;
        DrawingGridVerticalSpacingTwips = verticalSpacingTwips;
        DrawingGridHorizontalOriginTwips = horizontalOriginTwips;
        DrawingGridVerticalOriginTwips = verticalOriginTwips;
        DrawingGridHorizontalShow = horizontalShow;
        DrawingGridVerticalShow = verticalShow;
        SnapToDrawingGrid = snapToGrid;
        DrawingGridUsesMargins = useMargins;
        return this;
    }

    /// <summary>Sets the document default text direction.</summary>
    public RtfDocumentSettings SetDirection(RtfTextDirection? direction) {
        Direction = direction;
        return this;
    }

    internal bool HasAnyValue =>
        UnicodeSkipCount.HasValue ||
        DefaultFontId.HasValue ||
        DefaultTabWidthTwips.HasValue ||
        DefaultLanguageId.HasValue ||
        DefaultFarEastLanguageId.HasValue ||
        DefaultAlternateLanguageId.HasValue ||
        ViewKind.HasValue ||
        ViewScale.HasValue ||
        ZoomKind.HasValue ||
        ViewBackspaceBehavior.HasValue ||
        WidowOrphanControl.HasValue ||
        AutoHyphenation.HasValue ||
        HyphenateCaps.HasValue ||
        ConsecutiveHyphenLimit.HasValue ||
        HyphenationZoneTwips.HasValue ||
        FacingPages.HasValue ||
        MirrorMargins.HasValue ||
        FormProtection.HasValue ||
        RevisionProtection.HasValue ||
        AnnotationProtection.HasValue ||
        ReadOnlyProtection.HasValue ||
        TrackRevisions.HasValue ||
        RevisionDisplayStyle.HasValue ||
        RevisionBarPlacement.HasValue ||
        DrawingGridHorizontalSpacingTwips.HasValue ||
        DrawingGridVerticalSpacingTwips.HasValue ||
        DrawingGridHorizontalOriginTwips.HasValue ||
        DrawingGridVerticalOriginTwips.HasValue ||
        DrawingGridHorizontalShow.HasValue ||
        DrawingGridVerticalShow.HasValue ||
        SnapToDrawingGrid.HasValue ||
        DrawingGridUsesMargins.HasValue ||
        Direction.HasValue;

    private static void ValidateNonNegative(int? value, string parameterName) {
        if (value.HasValue && value.Value < 0) {
            throw new ArgumentOutOfRangeException(parameterName, "Document setting cannot be negative.");
        }
    }
}
