namespace OfficeIMO.OneNote;

/// <summary>
/// Logical kind of a typed OneNote page element.
/// </summary>
public enum OneNoteElementKind {
    /// <summary>Unknown or opaque content.</summary>
    Unknown = 0,
    /// <summary>An outline container.</summary>
    Outline = 1,
    /// <summary>A rich-text paragraph.</summary>
    Paragraph = 2,
    /// <summary>A table.</summary>
    Table = 3,
    /// <summary>An image.</summary>
    Image = 4,
    /// <summary>An embedded file.</summary>
    EmbeddedFile = 5,
    /// <summary>Ink or handwriting data.</summary>
    Ink = 6,
    /// <summary>Mathematical content.</summary>
    Math = 7,
    /// <summary>Audio or video recording content.</summary>
    Media = 8
}

/// <summary>
/// Base class for typed page elements.
/// </summary>
public abstract class OneNoteElement {
    /// <summary>Object identifier within the revision store. Serialization assigns and retains an identity for new content.</summary>
    public OneNoteExtendedGuid? Id { get; set; }

    /// <summary>Logical content kind.</summary>
    public abstract OneNoteElementKind Kind { get; }

    /// <summary>Optional absolute or parent-relative layout.</summary>
    public OneNoteLayout? Layout { get; set; }

    /// <summary>Author metadata associated with this element.</summary>
    public OneNoteAuthor? Author { get; set; }

    /// <summary>Note tags associated with this element.</summary>
    public IList<OneNoteTag> Tags { get; } = new List<OneNoteTag>();

    /// <summary>Unknown properties preserved in encoded source order.</summary>
    public IList<OneNoteOpaqueProperty> UnknownProperties { get; } = new List<OneNoteOpaqueProperty>();
}

/// <summary>
/// Placement and size of a OneNote element.
/// </summary>
public sealed class OneNoteLayout {
    /// <summary>Horizontal offset from the parent or page.</summary>
    public double? X { get; set; }

    /// <summary>Vertical offset from the parent or page.</summary>
    public double? Y { get; set; }

    /// <summary>Layout width.</summary>
    public double? Width { get; set; }

    /// <summary>Layout height.</summary>
    public double? Height { get; set; }

    /// <summary>Whether the element uses a tight layout.</summary>
    public bool? Tight { get; set; }

    /// <summary>Whether the element is placed in right-to-left reading order.</summary>
    public bool? RightToLeft { get; set; }

    /// <summary>Minimum outline width.</summary>
    public double? MinimumWidth { get; set; }

    /// <summary>Native alignment flags relative to the parent.</summary>
    public uint? AlignmentInParent { get; set; }

    /// <summary>Native self-alignment flags.</summary>
    public uint? AlignmentSelf { get; set; }

    /// <summary>Native collision-resolution priority.</summary>
    public uint? CollisionPriority { get; set; }

    /// <summary>Whether tight alignment is enabled.</summary>
    public bool? TightAlignment { get; set; }
}

/// <summary>
/// A freely positioned outline containing paragraphs, tables, and media.
/// </summary>
public sealed class OneNoteOutline : OneNoteElement {
    /// <inheritdoc />
    public override OneNoteElementKind Kind => OneNoteElementKind.Outline;

    /// <summary>Outline children in source order.</summary>
    public IList<OneNoteElement> Children { get; } = new List<OneNoteElement>();

    /// <summary>Marks a semantic container that must be emitted as an MS-ONE outline-element wrapper.</summary>
    internal bool IsOutlineElementWrapper { get; set; }

    /// <summary>List metadata carried by a preserved nonparagraph outline-element wrapper.</summary>
    internal OneNoteListInfo? WrapperList { get; set; }
}

/// <summary>
/// A rich-text paragraph and its nested outline elements.
/// </summary>
public sealed class OneNoteParagraph : OneNoteElement {
    /// <inheritdoc />
    public override OneNoteElementKind Kind => OneNoteElementKind.Paragraph;

    /// <summary>Formatted text runs in source order.</summary>
    public IList<OneNoteTextRun> Runs { get; } = new List<OneNoteTextRun>();

    /// <summary>Appends a structured inline mathematical expression.</summary>
    public OneNoteTextRun AddMath(OfficeIMO.Drawing.OfficeMathExpression expression) {
        var run = new OneNoteTextRun().SetMathExpression(expression);
        Runs.Add(run);
        return run;
    }

    /// <summary>Optional list marker metadata.</summary>
    public OneNoteListInfo? List { get; set; }

    /// <summary>Paragraph formatting.</summary>
    public OneNoteParagraphStyle Style { get; } = new OneNoteParagraphStyle();

    /// <summary>Nested outline elements under this paragraph.</summary>
    public IList<OneNoteElement> Children { get; } = new List<OneNoteElement>();

    internal OneNoteExtendedGuid? ContentObjectId { get; set; }
}

/// <summary>
/// A formatted text run.
/// </summary>
public sealed class OneNoteTextRun {
    internal OneNoteMathInlineDescriptor? MathDescriptor { get; set; }
    internal OfficeIMO.Drawing.OfficeMathExpression? PreservedMathExpression { get; set; }
    internal IReadOnlyList<OneNoteTextRun>? PreservedNativeMathRuns { get; set; }

    /// <summary>Unicode run text.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Optional structured inline mathematical expression. The reusable expression tree is owned by
    /// OfficeIMO.Drawing; OneNote maps it to and from the native rich-text math grammar.
    /// </summary>
    public OfficeIMO.Drawing.OfficeMathExpression? MathExpression { get; set; }

    /// <summary>Assigns a structured inline mathematical expression and its readable text projection.</summary>
    public OneNoteTextRun SetMathExpression(OfficeIMO.Drawing.OfficeMathExpression expression) {
        MathExpression = expression ?? throw new ArgumentNullException(nameof(expression));
        PreservedMathExpression = null;
        PreservedNativeMathRuns = null;
        Text = expression.ToPlainText();
        Style.IsMath = true;
        return this;
    }

    /// <summary>Run formatting.</summary>
    public OneNoteTextStyle Style { get; } = new OneNoteTextStyle();

    /// <summary>Optional hyperlink URI.</summary>
    public string? Hyperlink { get; set; }

    /// <summary>Whether the hyperlink is protected from editing.</summary>
    public bool HyperlinkProtected { get; set; }

    /// <summary>Unknown run properties preserved in encoded source order.</summary>
    public IList<OneNoteOpaqueProperty> UnknownProperties { get; } = new List<OneNoteOpaqueProperty>();

    internal OneNoteExtendedGuid? StyleObjectId { get; set; }
}

/// <summary>
/// Character formatting for a text run.
/// </summary>
public sealed class OneNoteTextStyle {
    /// <summary>Font family.</summary>
    public string? FontFamily { get; set; }

    /// <summary>Font size in points.</summary>
    public double? FontSize { get; set; }

    /// <summary>Font color encoded as ARGB.</summary>
    public uint? ColorArgb { get; set; }

    /// <summary>Highlight color encoded as ARGB.</summary>
    public uint? HighlightColorArgb { get; set; }

    /// <summary>Whether bold formatting is active.</summary>
    public bool? Bold { get; set; }

    /// <summary>Whether italic formatting is active.</summary>
    public bool? Italic { get; set; }

    /// <summary>Whether underline formatting is active.</summary>
    public bool? Underline { get; set; }

    /// <summary>Whether strikethrough formatting is active.</summary>
    public bool? Strikethrough { get; set; }

    /// <summary>Whether superscript formatting is active.</summary>
    public bool? Superscript { get; set; }

    /// <summary>Whether subscript formatting is active.</summary>
    public bool? Subscript { get; set; }

    /// <summary>LCID language identifier.</summary>
    public uint? LanguageId { get; set; }

    /// <summary>Whether OneNote marks this run as mathematical content.</summary>
    public bool? IsMath { get; set; }
}

/// <summary>
/// Paragraph alignment and spacing.
/// </summary>
public sealed class OneNoteParagraphStyle {
    internal OneNoteExtendedGuid? ObjectId { get; set; }

    /// <summary>Named paragraph style identifier.</summary>
    public string? StyleId { get; set; }

    /// <summary>Horizontal alignment.</summary>
    public OneNoteParagraphAlignment? Alignment { get; set; }

    /// <summary>Space before the paragraph in native half-inch units.</summary>
    public double? SpaceBefore { get; set; }

    /// <summary>Space after the paragraph in native half-inch units.</summary>
    public double? SpaceAfter { get; set; }

    /// <summary>Exact line spacing in native half-inch units.</summary>
    public double? ExactLineSpacing { get; set; }
}

/// <summary>
/// Paragraph horizontal alignment.
/// </summary>
public enum OneNoteParagraphAlignment {
    /// <summary>Left aligned.</summary>
    Left = 0,
    /// <summary>Centered.</summary>
    Center = 1,
    /// <summary>Right aligned.</summary>
    Right = 2,
    /// <summary>Justified.</summary>
    Justify = 3
}

/// <summary>
/// List marker metadata for a paragraph.
/// </summary>
public sealed class OneNoteListInfo {
    /// <summary>Largest zero-based list level representable by the native one-byte child-level property.</summary>
    public const int MaxLevel = byte.MaxValue - 1;

    internal OneNoteExtendedGuid? ObjectId { get; set; }

    /// <summary>Whether the list is ordered.</summary>
    public bool Ordered { get; set; }

    /// <summary>MS-ONE number-list format value.</summary>
    public uint? Format { get; set; }

    /// <summary>Zero-based nesting level. Native writing accepts values from 0 through <see cref="MaxLevel"/>.</summary>
    public int Level { get; set; }

    /// <summary>Whether numbering restarts at this item.</summary>
    public bool Restart { get; set; }

    /// <summary>Displayed list index when available.</summary>
    public int? DisplayIndex { get; set; }

    /// <summary>Bullet or number font.</summary>
    public string? FontFamily { get; set; }
}

/// <summary>
/// A OneNote table.
/// </summary>
public sealed class OneNoteTable : OneNoteElement {
    /// <inheritdoc />
    public override OneNoteElementKind Kind => OneNoteElementKind.Table;

    /// <summary>Whether table borders are visible.</summary>
    public bool BordersVisible { get; set; }

    /// <summary>Column widths in OneNote layout units.</summary>
    public IList<double> ColumnWidths { get; } = new List<double>();

    /// <summary>Table rows in source order.</summary>
    public IList<OneNoteTableRow> Rows { get; } = new List<OneNoteTableRow>();
}

/// <summary>
/// A row in a OneNote table.
/// </summary>
public sealed class OneNoteTableRow {
    internal OneNoteExtendedGuid? ObjectId { get; set; }

    /// <summary>Cells in source order.</summary>
    public IList<OneNoteTableCell> Cells { get; } = new List<OneNoteTableCell>();
}

/// <summary>
/// A cell in a OneNote table.
/// </summary>
public sealed class OneNoteTableCell {
    internal OneNoteExtendedGuid? ObjectId { get; set; }

    /// <summary>Cell shading color encoded as ARGB.</summary>
    public uint? ShadingColorArgb { get; set; }

    /// <summary>Cell content in source order.</summary>
    public IList<OneNoteElement> Content { get; } = new List<OneNoteElement>();

    /// <summary>Unknown cell properties preserved in encoded source order.</summary>
    public IList<OneNoteOpaqueProperty> UnknownProperties { get; } = new List<OneNoteOpaqueProperty>();
}
