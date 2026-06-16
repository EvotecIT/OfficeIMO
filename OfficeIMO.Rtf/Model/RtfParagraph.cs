namespace OfficeIMO.Rtf;

/// <summary>
/// Paragraph in an RTF document.
/// </summary>
public sealed class RtfParagraph : IRtfBlock {
    private readonly List<RtfRun> _runs = new List<RtfRun>();
    private readonly List<IRtfInline> _inlines = new List<IRtfInline>();
    private readonly List<RtfTabStop> _tabStops = new List<RtfTabStop>();

    /// <summary>Paragraph runs.</summary>
    public IReadOnlyList<RtfRun> Runs => _runs.AsReadOnly();

    /// <summary>Paragraph inline content in original order, including non-text markers such as bookmarks.</summary>
    public IReadOnlyList<IRtfInline> Inlines => _inlines.AsReadOnly();

    /// <summary>Paragraph tab stops in RTF declaration order.</summary>
    public IReadOnlyList<RtfTabStop> TabStops => _tabStops.AsReadOnly();

    /// <summary>Paragraph alignment.</summary>
    public RtfTextAlignment Alignment { get; set; } = RtfTextAlignment.Left;

    /// <summary>Explicit paragraph text direction represented by <c>\ltrpar</c> or <c>\rtlpar</c>.</summary>
    public RtfTextDirection? Direction { get; set; }

    /// <summary>Optional paragraph stylesheet id.</summary>
    public int? StyleId { get; set; }

    /// <summary>Optional list id for paragraphs bound to an RTF list override.</summary>
    public int? ListId { get; set; }

    /// <summary>Optional list definition id resolved through the list override table.</summary>
    public int? ListDefinitionId { get; set; }

    /// <summary>Optional zero-based list level.</summary>
    public int? ListLevel { get; set; }

    /// <summary>Basic list marker kind.</summary>
    public RtfListKind ListKind { get; set; } = RtfListKind.None;

    /// <summary>Word 6/95 legacy paragraph numbering metadata from <c>\pn*</c> controls.</summary>
    public RtfLegacyNumbering LegacyNumbering { get; } = new RtfLegacyNumbering();

    /// <summary>Word list marker fallback text represented by a paragraph-level <c>\listtext</c> destination.</summary>
    public RtfParagraph? ListText { get; private set; }

    /// <summary>Left indentation in twips.</summary>
    public int? LeftIndentTwips { get; set; }

    /// <summary>Right indentation in twips.</summary>
    public int? RightIndentTwips { get; set; }

    /// <summary>First-line indentation in twips.</summary>
    public int? FirstLineIndentTwips { get; set; }

    /// <summary>Space before paragraph in twips.</summary>
    public int? SpaceBeforeTwips { get; set; }

    /// <summary>Space after paragraph in twips.</summary>
    public int? SpaceAfterTwips { get; set; }

    /// <summary>Whether space before the paragraph is automatically determined by the renderer.</summary>
    public bool? SpaceBeforeAuto { get; set; }

    /// <summary>Whether space after the paragraph is automatically determined by the renderer.</summary>
    public bool? SpaceAfterAuto { get; set; }

    /// <summary>Raw RTF line spacing value from <c>\sl</c>.</summary>
    public int? LineSpacingTwips { get; set; }

    /// <summary>Whether <c>\sl</c> is interpreted as multiple line spacing by <c>\slmult</c>. Null means not specified.</summary>
    public bool? LineSpacingMultiple { get; set; }

    /// <summary>One-based color table index used by paragraph background shading.</summary>
    public int? BackgroundColorIndex { get; set; }

    /// <summary>One-based color table index used by paragraph pattern foreground shading.</summary>
    public int? ShadingForegroundColorIndex { get; set; }

    /// <summary>Raw RTF <c>\shading</c> value, where 10000 represents 100 percent.</summary>
    public int? ShadingPatternPercent { get; set; }

    /// <summary>Named RTF shading pattern.</summary>
    public RtfShadingPattern ShadingPattern { get; set; } = RtfShadingPattern.None;

    /// <summary>Top paragraph border.</summary>
    public RtfParagraphBorder TopBorder { get; } = new RtfParagraphBorder();

    /// <summary>Left paragraph border.</summary>
    public RtfParagraphBorder LeftBorder { get; } = new RtfParagraphBorder();

    /// <summary>Bottom paragraph border.</summary>
    public RtfParagraphBorder BottomBorder { get; } = new RtfParagraphBorder();

    /// <summary>Right paragraph border.</summary>
    public RtfParagraphBorder RightBorder { get; } = new RtfParagraphBorder();

    /// <summary>Whether a page break should be emitted before this paragraph.</summary>
    public bool PageBreakBefore { get; set; }

    /// <summary>Whether this paragraph should stay on the same page as the following paragraph.</summary>
    public bool KeepWithNext { get; set; }

    /// <summary>Whether all lines in this paragraph should stay together on the same page.</summary>
    public bool KeepLinesTogether { get; set; }

    /// <summary>Whether line numbering is suppressed for this paragraph.</summary>
    public bool SuppressLineNumbers { get; set; }

    /// <summary>Paragraph-level automatic hyphenation override represented by <c>\hyphpar</c>.</summary>
    public bool? AutoHyphenation { get; set; }

    /// <summary>Whether contextual paragraph spacing is explicitly enabled or disabled.</summary>
    public bool? ContextualSpacing { get; set; }

    /// <summary>Whether the right indent is automatically adjusted when a document grid is defined.</summary>
    public bool? AdjustRightIndent { get; set; }

    /// <summary>Whether paragraph lines snap to the document line grid.</summary>
    public bool? SnapToLineGrid { get; set; }

    /// <summary>Explicit paragraph widow/orphan control. Null means not specified.</summary>
    public bool? WidowControl { get; set; }

    /// <summary>Optional paragraph outline level.</summary>
    public int? OutlineLevel { get; set; }

    /// <summary>Paragraph revision save identifier represented by <c>\pararsid</c>.</summary>
    public int? RevisionSaveId { get; set; }

    /// <summary>Absolute positioning, text wrapping, and drop-cap metadata for this paragraph.</summary>
    public RtfParagraphFrame Frame { get; } = new RtfParagraphFrame();

    /// <summary>Adds a text run.</summary>
    public RtfRun AddText(string text) {
        var run = new RtfRun(text);
        _runs.Add(run);
        _inlines.Add(run);
        return run;
    }

    /// <summary>Adds a bookmark start marker at the current paragraph position.</summary>
    public RtfBookmarkMarker AddBookmarkStart(string name) {
        return AddBookmarkMarker(RtfBookmarkMarkerKind.Start, name);
    }

    /// <summary>Adds a bookmark end marker at the current paragraph position.</summary>
    public RtfBookmarkMarker AddBookmarkEnd(string name) {
        return AddBookmarkMarker(RtfBookmarkMarkerKind.End, name);
    }

    /// <summary>Adds an RTF field at the current paragraph position.</summary>
    public RtfField AddField(string instruction) {
        var field = new RtfField(instruction);
        AddField(field);
        return field;
    }

    /// <summary>Adds a generated text marker at the current paragraph position.</summary>
    public RtfGeneratedText AddGeneratedText(RtfGeneratedTextKind kind, string? fallbackText = null) {
        var generatedText = new RtfGeneratedText(kind, fallbackText);
        AddGeneratedText(generatedText);
        return generatedText;
    }

    /// <summary>Adds a generated current page number marker.</summary>
    public RtfGeneratedText AddPageNumber(string? fallbackText = null) {
        return AddGeneratedText(RtfGeneratedTextKind.PageNumber, fallbackText);
    }

    /// <summary>Adds a generated current section number marker.</summary>
    public RtfGeneratedText AddSectionNumber(string? fallbackText = null) {
        return AddGeneratedText(RtfGeneratedTextKind.SectionNumber, fallbackText);
    }

    /// <summary>Adds a generated current date marker.</summary>
    public RtfGeneratedText AddCurrentDate(string? fallbackText = null) {
        return AddGeneratedText(RtfGeneratedTextKind.CurrentDate, fallbackText);
    }

    /// <summary>Adds a generated current date marker in long format.</summary>
    public RtfGeneratedText AddCurrentDateLong(string? fallbackText = null) {
        return AddGeneratedText(RtfGeneratedTextKind.CurrentDateLong, fallbackText);
    }

    /// <summary>Adds a generated current date marker in abbreviated format.</summary>
    public RtfGeneratedText AddCurrentDateAbbreviated(string? fallbackText = null) {
        return AddGeneratedText(RtfGeneratedTextKind.CurrentDateAbbreviated, fallbackText);
    }

    /// <summary>Adds a generated current time marker.</summary>
    public RtfGeneratedText AddCurrentTime(string? fallbackText = null) {
        return AddGeneratedText(RtfGeneratedTextKind.CurrentTime, fallbackText);
    }

    /// <summary>Adds an automatic note reference marker and attaches the supplied note to it.</summary>
    public RtfGeneratedText AddNoteReference(RtfNote note, string? fallbackText = null) {
        if (note == null) throw new ArgumentNullException(nameof(note));
        RtfGeneratedText generatedText = AddGeneratedText(RtfGeneratedTextKind.NoteReference, fallbackText);
        generatedText.Note = note;
        return generatedText;
    }

    /// <summary>Adds an embedded or linked object at the current paragraph position.</summary>
    public RtfObject AddObject(RtfObjectKind kind = RtfObjectKind.Unknown, byte[]? data = null) {
        var rtfObject = new RtfObject(kind, data);
        AddObject(rtfObject);
        return rtfObject;
    }

    /// <summary>Adds a drawing shape at the current paragraph position.</summary>
    public RtfShape AddShape() {
        var shape = new RtfShape();
        AddShape(shape);
        return shape;
    }

    /// <summary>Adds a picture at the current paragraph position.</summary>
    public RtfImage AddImage(RtfImageFormat format, byte[] data) {
        var image = new RtfImage(format, data);
        AddImage(image);
        return image;
    }

    /// <summary>Adds a superscript footnote reference run with note text.</summary>
    public RtfRun AddFootnote(string referenceText, string noteText) {
        return AddNoteReference(referenceText, RtfNoteKind.Footnote, noteText);
    }

    /// <summary>Adds a superscript endnote reference run with note text.</summary>
    public RtfRun AddEndnote(string referenceText, string noteText) {
        return AddNoteReference(referenceText, RtfNoteKind.Endnote, noteText);
    }

    /// <summary>Adds an annotation reference run with note text.</summary>
    public RtfRun AddAnnotation(string referenceText, string noteText) {
        return AddNoteReference(referenceText, RtfNoteKind.Annotation, noteText);
    }

    /// <summary>Adds a line break at the current paragraph position.</summary>
    public RtfBreak AddLineBreak() {
        return AddBreak(RtfBreakKind.Line);
    }

    /// <summary>Adds a soft line break at the current paragraph position.</summary>
    public RtfBreak AddSoftLineBreak() {
        return AddBreak(RtfBreakKind.SoftLine);
    }

    /// <summary>Adds a page break at the current paragraph position.</summary>
    public RtfBreak AddPageBreak() {
        return AddBreak(RtfBreakKind.Page);
    }

    /// <summary>Adds a soft page break at the current paragraph position.</summary>
    public RtfBreak AddSoftPageBreak() {
        return AddBreak(RtfBreakKind.SoftPage);
    }

    /// <summary>Adds a column break at the current paragraph position.</summary>
    public RtfBreak AddColumnBreak() {
        return AddBreak(RtfBreakKind.Column);
    }

    /// <summary>Adds a tab stop to the paragraph.</summary>
    public RtfTabStop AddTabStop(int positionTwips, RtfTabAlignment alignment = RtfTabAlignment.Left, RtfTabLeader leader = RtfTabLeader.None) {
        var tabStop = new RtfTabStop(positionTwips, alignment, leader);
        _tabStops.Add(tabStop);
        return tabStop;
    }

    internal void AddRun(RtfRun run) {
        run = run ?? throw new ArgumentNullException(nameof(run));
        _runs.Add(run);
        _inlines.Add(run);
    }

    internal void AddBookmarkMarker(RtfBookmarkMarker marker) {
        _inlines.Add(marker ?? throw new ArgumentNullException(nameof(marker)));
    }

    internal void AddField(RtfField field) {
        _inlines.Add(field ?? throw new ArgumentNullException(nameof(field)));
    }

    internal void AddGeneratedText(RtfGeneratedText generatedText) {
        _inlines.Add(generatedText ?? throw new ArgumentNullException(nameof(generatedText)));
    }

    internal void AddObject(RtfObject rtfObject) {
        _inlines.Add(rtfObject ?? throw new ArgumentNullException(nameof(rtfObject)));
    }

    internal void AddShape(RtfShape shape) {
        _inlines.Add(shape ?? throw new ArgumentNullException(nameof(shape)));
    }

    internal void AddImage(RtfImage image) {
        _inlines.Add(image ?? throw new ArgumentNullException(nameof(image)));
    }

    /// <summary>Adds an explicit break at the current paragraph position.</summary>
    public RtfBreak AddBreak(RtfBreakKind kind) {
        var rtfBreak = new RtfBreak(kind);
        _inlines.Add(rtfBreak);
        return rtfBreak;
    }

    internal void AddInline(IRtfInline inline) {
        switch (inline) {
            case RtfRun run:
                AddRun(run);
                break;
            case RtfBookmarkMarker marker:
                AddBookmarkMarker(marker);
                break;
            case RtfField field:
                AddField(field);
                break;
            case RtfGeneratedText generatedText:
                AddGeneratedText(new RtfGeneratedText(generatedText.Kind, generatedText.FallbackText) {
                    Note = generatedText.Note
                });
                break;
            case RtfObject rtfObject:
                AddObject(rtfObject);
                break;
            case RtfShape shape:
                AddShape(shape);
                break;
            case RtfImage image:
                AddImage(image);
                break;
            case RtfBreak rtfBreak:
                AddBreak(rtfBreak.Kind);
                break;
            default:
                throw new ArgumentException("Unsupported RTF inline type.", nameof(inline));
        }
    }

    internal void ReplaceTabStops(IEnumerable<RtfTabStop> tabStops) {
        _tabStops.Clear();
        foreach (RtfTabStop tabStop in tabStops) {
            _tabStops.Add(new RtfTabStop(tabStop.PositionTwips, tabStop.Alignment, tabStop.Leader));
        }
    }

    /// <summary>Sets paragraph alignment.</summary>
    public RtfParagraph SetAlignment(RtfTextAlignment alignment) {
        Alignment = alignment;
        return this;
    }

    /// <summary>Sets explicit paragraph text direction.</summary>
    public RtfParagraph SetDirection(RtfTextDirection? direction) {
        Direction = direction;
        return this;
    }

    /// <summary>Sets the paragraph style id.</summary>
    public RtfParagraph SetStyle(int styleId) {
        StyleId = styleId;
        return this;
    }

    /// <summary>Marks the paragraph as a list item.</summary>
    public RtfParagraph SetList(int listId = 1, int level = 0, RtfListKind kind = RtfListKind.Bullet) {
        if (level < 0) throw new ArgumentOutOfRangeException(nameof(level), "List level cannot be negative.");
        ListId = listId;
        ListLevel = level;
        ListKind = kind;
        return this;
    }

    /// <summary>Configures Word 6/95 legacy paragraph numbering metadata.</summary>
    public RtfParagraph SetLegacyNumbering(Action<RtfLegacyNumbering> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        LegacyNumbering.Enabled = true;
        configure(LegacyNumbering);
        return this;
    }

    /// <summary>Sets plain fallback list marker text emitted in a <c>\listtext</c> destination.</summary>
    public RtfParagraph SetListText(string text) {
        var listText = new RtfParagraph();
        listText.AddText(text ?? string.Empty);
        ListText = listText;
        return this;
    }

    /// <summary>Configures rich fallback list marker content emitted in a <c>\listtext</c> destination.</summary>
    public RtfParagraph SetListText(Action<RtfParagraph> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        var listText = new RtfParagraph();
        configure(listText);
        ListText = listText;
        return this;
    }

    /// <summary>Clears fallback list marker text.</summary>
    public RtfParagraph ClearListText() {
        ListText = null;
        return this;
    }

    internal void SetParsedListText(RtfParagraph? listText) {
        ListText = listText;
    }

    /// <summary>Sets paragraph indentation in twips.</summary>
    public RtfParagraph SetIndentation(int? leftTwips = null, int? rightTwips = null, int? firstLineTwips = null) {
        LeftIndentTwips = leftTwips;
        RightIndentTwips = rightTwips;
        FirstLineIndentTwips = firstLineTwips;
        return this;
    }

    /// <summary>Sets raw RTF paragraph line spacing controls.</summary>
    public RtfParagraph SetLineSpacing(int? lineSpacingTwips, bool? multiple = null) {
        LineSpacingTwips = lineSpacingTwips;
        LineSpacingMultiple = multiple;
        return this;
    }

    /// <summary>Sets paragraph spacing before and after this paragraph.</summary>
    public RtfParagraph SetParagraphSpacing(int? beforeTwips = null, int? afterTwips = null, bool? beforeAuto = null, bool? afterAuto = null) {
        SpaceBeforeTwips = beforeTwips;
        SpaceAfterTwips = afterTwips;
        SpaceBeforeAuto = beforeAuto;
        SpaceAfterAuto = afterAuto;
        return this;
    }

    /// <summary>Sets contextual paragraph spacing.</summary>
    public RtfParagraph SetContextualSpacing(bool? enabled = true) {
        ContextualSpacing = enabled;
        return this;
    }

    /// <summary>Sets automatic right-indent adjustment for document-grid layouts.</summary>
    public RtfParagraph SetAdjustRightIndent(bool? enabled = true) {
        AdjustRightIndent = enabled;
        return this;
    }

    /// <summary>Sets whether paragraph lines snap to the document line grid.</summary>
    public RtfParagraph SetSnapToLineGrid(bool? enabled = true) {
        SnapToLineGrid = enabled;
        return this;
    }

    /// <summary>Sets paragraph background shading to a one-based color table index.</summary>
    public RtfParagraph SetBackgroundColor(int? colorIndex) {
        BackgroundColorIndex = colorIndex;
        return this;
    }

    /// <summary>Sets paragraph shading color and pattern metadata.</summary>
    public RtfParagraph SetShading(int? backgroundColorIndex, int? foregroundColorIndex = null, int? patternPercent = null, RtfShadingPattern pattern = RtfShadingPattern.None) {
        BackgroundColorIndex = backgroundColorIndex;
        ShadingForegroundColorIndex = foregroundColorIndex;
        ShadingPatternPercent = patternPercent;
        ShadingPattern = pattern;
        return this;
    }

    /// <summary>Sets paragraph pagination controls.</summary>
    public RtfParagraph SetPagination(bool? pageBreakBefore = null, bool? keepWithNext = null, bool? keepLinesTogether = null, bool? widowControl = null, bool? suppressLineNumbers = null, bool? autoHyphenation = null) {
        if (pageBreakBefore.HasValue) {
            PageBreakBefore = pageBreakBefore.Value;
        }

        if (keepWithNext.HasValue) {
            KeepWithNext = keepWithNext.Value;
        }

        if (keepLinesTogether.HasValue) {
            KeepLinesTogether = keepLinesTogether.Value;
        }

        if (suppressLineNumbers.HasValue) {
            SuppressLineNumbers = suppressLineNumbers.Value;
        }

        AutoHyphenation = autoHyphenation;
        WidowControl = widowControl;
        return this;
    }

    /// <summary>Sets the paragraph outline level.</summary>
    public RtfParagraph SetOutlineLevel(int? outlineLevel) {
        OutlineLevel = outlineLevel;
        return this;
    }

    /// <summary>Sets the paragraph revision save identifier represented by <c>\pararsid</c>.</summary>
    public RtfParagraph SetRevisionSaveId(int? id) {
        if (id.HasValue && id.Value < 0) throw new ArgumentOutOfRangeException(nameof(id), "Paragraph revision save id cannot be negative.");
        RevisionSaveId = id;
        return this;
    }

    /// <summary>Configures absolute positioning frame metadata for this paragraph.</summary>
    public RtfParagraph SetFrame(Action<RtfParagraphFrame> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        configure(Frame);
        return this;
    }

    /// <summary>Sets paragraph border formatting for one side.</summary>
    public RtfParagraph SetBorder(RtfParagraphBorderSide side, RtfParagraphBorderStyle style, int? width = null, int? colorIndex = null) {
        RtfParagraphBorder border = GetBorder(side);
        border.Style = style;
        border.Width = width;
        border.ColorIndex = colorIndex;
        return this;
    }

    /// <summary>Returns plain text for the paragraph.</summary>
    public string ToPlainText() {
        var builder = new StringBuilder();
        foreach (IRtfInline inline in _inlines) {
            switch (inline) {
                case RtfRun run:
                    builder.Append(run.Text);
                    break;
                case RtfField field:
                    builder.Append(field.ToPlainText());
                    break;
                case RtfGeneratedText generatedText:
                    builder.Append(generatedText.ToPlainText());
                    break;
                case RtfObject rtfObject:
                    builder.Append(rtfObject.ToPlainText());
                    break;
                case RtfShape shape:
                    builder.Append(shape.ToPlainText());
                    break;
                case RtfBreak rtfBreak:
                    builder.Append(GetBreakText(rtfBreak.Kind));
                    break;
            }
        }

        return builder.ToString();
    }

    private static string GetBreakText(RtfBreakKind kind) {
        switch (kind) {
            case RtfBreakKind.SoftPage:
            case RtfBreakKind.Page:
                return "\f";
            case RtfBreakKind.Column:
                return "\v";
            case RtfBreakKind.SoftLine:
            default:
                return Environment.NewLine;
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

    private RtfRun AddNoteReference(string referenceText, RtfNoteKind kind, string noteText) {
        var note = new RtfNote(kind);
        note.AddParagraph(noteText);
        RtfRun run = AddText(referenceText);
        run.Note = note;
        if (kind == RtfNoteKind.Footnote || kind == RtfNoteKind.Endnote) {
            run.VerticalPosition = RtfVerticalPosition.Superscript;
        }

        return run;
    }

    private RtfBookmarkMarker AddBookmarkMarker(RtfBookmarkMarkerKind kind, string name) {
        var marker = new RtfBookmarkMarker(kind, name);
        _inlines.Add(marker);
        return marker;
    }
}
