namespace OfficeIMO.Rtf;

/// <summary>
/// Semantic RTF section containing ordered document blocks and section-level page layout.
/// </summary>
public sealed partial class RtfSection {
    private readonly List<IRtfBlock> _blocks = new List<IRtfBlock>();
    private readonly List<RtfSectionColumn> _columns = new List<RtfSectionColumn>();
    private readonly RtfDocument? _document;

    internal RtfSection(RtfDocument? document = null) {
        _document = document;
    }

    /// <summary>Ordered blocks contained by the section.</summary>
    public IReadOnlyList<IRtfBlock> Blocks => _blocks.AsReadOnly();

    /// <summary>Section page size, margins, and orientation.</summary>
    public RtfPageSetup PageSetup { get; } = new RtfPageSetup();

    /// <summary>Section-level footnote and endnote numbering settings.</summary>
    public RtfNoteSettings NoteSettings { get; } = new RtfNoteSettings();

    /// <summary>Section break behavior.</summary>
    public RtfSectionBreakKind BreakKind { get; set; } = RtfSectionBreakKind.NextPage;

    /// <summary>Number of text columns in the section.</summary>
    public int? ColumnCount { get; set; }

    /// <summary>Space between section columns in twips.</summary>
    public int? ColumnSpaceTwips { get; set; }

    /// <summary>Whether a vertical separator line is shown between columns.</summary>
    public bool ColumnSeparator { get; set; }

    /// <summary>Unequal section column definitions in section order.</summary>
    public IReadOnlyList<RtfSectionColumn> Columns => _columns.AsReadOnly();

    /// <summary>Section line-numbering settings.</summary>
    public RtfLineNumbering LineNumbering { get; } = new RtfLineNumbering();

    /// <summary>Vertical text alignment for pages in the section.</summary>
    public RtfSectionVerticalAlignment? VerticalAlignment { get; set; }

    /// <summary>Section text direction represented by <c>\ltrsect</c> or <c>\rtlsect</c>.</summary>
    public RtfTextDirection? Direction { get; set; }

    /// <summary>Sets vertical text alignment for pages in the section.</summary>
    public RtfSection SetVerticalAlignment(RtfSectionVerticalAlignment? alignment) {
        VerticalAlignment = alignment;
        return this;
    }

    /// <summary>Sets the section text direction.</summary>
    public RtfSection SetDirection(RtfTextDirection? direction) {
        Direction = direction;
        return this;
    }

    /// <summary>Adds an unequal section column definition.</summary>
    public RtfSectionColumn AddColumn(int? widthTwips = null, int? spaceAfterTwips = null) {
        var column = new RtfSectionColumn(widthTwips, spaceAfterTwips);
        _columns.Add(column);
        return column;
    }

    /// <summary>Clears unequal section column definitions.</summary>
    public RtfSection ClearColumns() {
        _columns.Clear();
        return this;
    }

    /// <summary>Adds a paragraph to the section.</summary>
    public RtfParagraph AddParagraph(string? text = null) {
        var paragraph = new RtfParagraph();
        if (!string.IsNullOrEmpty(text)) {
            paragraph.AddText(text!);
        }

        _blocks.Add(paragraph);
        _document?.AddParsedBlock(paragraph);
        return paragraph;
    }

    /// <summary>Adds a table block to the section.</summary>
    public RtfTable AddTable(int rows, int columns) {
        if (rows < 0) throw new ArgumentOutOfRangeException(nameof(rows), "Row count cannot be negative.");
        if (columns <= 0) throw new ArgumentOutOfRangeException(nameof(columns), "Column count must be greater than zero.");

        var table = new RtfTable();
        const int defaultColumnWidthTwips = 2400;
        for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
            RtfTableRow row = table.AddRow();
            for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
                row.AddCell((columnIndex + 1) * defaultColumnWidthTwips);
            }
        }

        _blocks.Add(table);
        _document?.AddParsedBlock(table);
        return table;
    }

    /// <summary>Adds a picture block to the section.</summary>
    public RtfImage AddImage(RtfImageFormat format, byte[] data) {
        var image = new RtfImage(format, data);
        _blocks.Add(image);
        _document?.AddParsedBlock(image);
        return image;
    }

    internal bool HasAnyLayoutValue =>
        PageSetup.HasAnyValue ||
        NoteSettings.HasAnyValue ||
        BreakKind != RtfSectionBreakKind.NextPage ||
        ColumnCount.HasValue ||
        ColumnSpaceTwips.HasValue ||
        ColumnSeparator ||
        _columns.Any(column => column.HasAnyValue) ||
        VerticalAlignment.HasValue ||
        Direction.HasValue ||
        LineNumbering.HasAnyValue;

    internal void AddParsedBlock(IRtfBlock block) {
        _blocks.Add(block ?? throw new ArgumentNullException(nameof(block)));
    }

    internal void ResetLayout() {
        PageSetup.Clear();
        NoteSettings.Clear();
        BreakKind = RtfSectionBreakKind.NextPage;
        ColumnCount = null;
        ColumnSpaceTwips = null;
        ColumnSeparator = false;
        _columns.Clear();
        VerticalAlignment = null;
        Direction = null;
        LineNumbering.Clear();
    }

    internal RtfSectionColumn EnsureColumn(int oneBasedIndex) {
        if (oneBasedIndex <= 0) {
            oneBasedIndex = _columns.Count + 1;
        }

        while (_columns.Count < oneBasedIndex) {
            _columns.Add(new RtfSectionColumn());
        }

        return _columns[oneBasedIndex - 1];
    }
}
