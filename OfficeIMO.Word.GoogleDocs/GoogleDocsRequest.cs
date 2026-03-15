namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Base request emitted by the Google Docs batch compiler.
    /// </summary>
    public abstract class GoogleDocsRequest {
        protected GoogleDocsRequest(string kind) {
            Kind = kind ?? throw new ArgumentNullException(nameof(kind));
        }

        public string Kind { get; }
        public int SectionIndex { get; set; }
        public int ElementIndex { get; set; }
        public GoogleDocsSectionStyle? SectionStyle { get; set; }
    }

    /// <summary>
    /// Inserts one paragraph block into the target Google Doc.
    /// </summary>
    public sealed class GoogleDocsInsertParagraphRequest : GoogleDocsRequest {
        public GoogleDocsInsertParagraphRequest() : base("insertParagraph") {
        }

        public GoogleDocsParagraph Paragraph { get; set; } = new GoogleDocsParagraph();
    }

    /// <summary>
    /// Inserts one table block into the target Google Doc.
    /// </summary>
    public sealed class GoogleDocsInsertTableRequest : GoogleDocsRequest {
        public GoogleDocsInsertTableRequest() : base("insertTable") {
        }

        public GoogleDocsTable Table { get; set; } = new GoogleDocsTable();
        public bool StartsNewSectionBefore { get; set; }
        public string? SectionBreakType { get; set; }
    }

    /// <summary>
    /// Normalized header/footer segment payload compiled from OfficeIMO snapshot data.
    /// </summary>
    public sealed class GoogleDocsSegment {
        private readonly List<GoogleDocsRequest> _requests = new List<GoogleDocsRequest>();

        public int SectionIndex { get; set; }
        public string Kind { get; set; } = string.Empty;
        public string Variant { get; set; } = string.Empty;
        public int TableCount { get; set; }
        public IReadOnlyList<GoogleDocsRequest> Requests => _requests;
        public IReadOnlyList<GoogleDocsParagraph> Paragraphs => _requests
            .OfType<GoogleDocsInsertParagraphRequest>()
            .Select(request => request.Paragraph)
            .ToList();
        public IReadOnlyList<GoogleDocsTable> Tables => _requests
            .OfType<GoogleDocsInsertTableRequest>()
            .Select(request => request.Table)
            .ToList();

        internal void AddRequest(GoogleDocsRequest request) {
            if (request == null) throw new ArgumentNullException(nameof(request));
            _requests.Add(request);
        }
    }

    /// <summary>
    /// Normalized paragraph payload used by the compiler output.
    /// </summary>
    public sealed class GoogleDocsParagraph {
        private readonly List<GoogleDocsParagraphRun> _runs = new List<GoogleDocsParagraphRun>();
        private readonly List<GoogleDocsTabStop> _tabStops = new List<GoogleDocsTabStop>();

        public string Text { get; set; } = string.Empty;
        public string? StyleId { get; set; }
        public string? StyleName { get; set; }
        public bool StartsNewSectionBefore { get; set; }
        public string? SectionBreakType { get; set; }
        public bool IsListItem { get; set; }
        public bool? IsOrderedList { get; set; }
        public int? ListLevel { get; set; }
        public string? ListStyleName { get; set; }
        public string? Alignment { get; set; }
        public double? IndentStartPoints { get; set; }
        public double? IndentEndPoints { get; set; }
        public double? IndentFirstLinePoints { get; set; }
        public double? SpaceAbovePoints { get; set; }
        public double? SpaceBelowPoints { get; set; }
        public double? LineSpacingPercent { get; set; }
        public string? ShadingFillColorHex { get; set; }
        public GoogleDocsParagraphBorder? LeftBorder { get; set; }
        public GoogleDocsParagraphBorder? RightBorder { get; set; }
        public GoogleDocsParagraphBorder? TopBorder { get; set; }
        public GoogleDocsParagraphBorder? BottomBorder { get; set; }
        public bool IsRightToLeft { get; set; }
        public bool KeepWithNext { get; set; }
        public bool KeepLinesTogether { get; set; }
        public bool AvoidWidowAndOrphan { get; set; }
        public bool PageBreakBefore { get; set; }
        public string? BookmarkName { get; set; }
        public int? BookmarkId { get; set; }
        public IReadOnlyList<GoogleDocsParagraphRun> Runs => _runs;
        public IReadOnlyList<GoogleDocsTabStop> TabStops => _tabStops;

        internal void AddRun(GoogleDocsParagraphRun run) {
            if (run == null) throw new ArgumentNullException(nameof(run));
            _runs.Add(run);
        }

        internal void AddTabStop(GoogleDocsTabStop tabStop) {
            if (tabStop == null) throw new ArgumentNullException(nameof(tabStop));
            _tabStops.Add(tabStop);
        }
    }

    /// <summary>
    /// Normalized run payload used by the compiler output.
    /// </summary>
    public sealed class GoogleDocsParagraphRun {
        public string Text { get; set; } = string.Empty;
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strike { get; set; }
        public int? FontSize { get; set; }
        public string? FontFamily { get; set; }
        public string? ForegroundColorHex { get; set; }
        public string? HighlightColor { get; set; }
        public string? VerticalTextAlignment { get; set; }
        public string? CapsStyle { get; set; }
        public GoogleDocsLink? Link { get; set; }
        public GoogleDocsFootnote? Footnote { get; set; }
        public GoogleDocsInlineImage? InlineImage { get; set; }
    }

    /// <summary>
    /// Normalized hyperlink payload used by the compiler output.
    /// </summary>
    public sealed class GoogleDocsLink {
        public string? Uri { get; set; }
        public string? Anchor { get; set; }
    }

    /// <summary>
    /// Normalized inline image payload used by the compiler output.
    /// </summary>
    public sealed class GoogleDocsInlineImage {
        public string? FilePath { get; set; }
        public string? FileName { get; set; }
        public string? ContentType { get; set; }
        public byte[]? Bytes { get; set; }
        public string? Description { get; set; }
        public string? Title { get; set; }
        public double? Width { get; set; }
        public double? Height { get; set; }
        public bool IsInline { get; set; }
        public string? WrapText { get; set; }
    }

    /// <summary>
    /// Normalized footnote payload used by the compiler output.
    /// </summary>
    public sealed class GoogleDocsFootnote {
        private readonly List<GoogleDocsParagraph> _paragraphs = new List<GoogleDocsParagraph>();

        public long? ReferenceId { get; set; }
        public IReadOnlyList<GoogleDocsParagraph> Paragraphs => _paragraphs;

        internal void AddParagraph(GoogleDocsParagraph paragraph) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            _paragraphs.Add(paragraph);
        }
    }

    /// <summary>
    /// Normalized table payload used by the compiler output.
    /// </summary>
    public sealed class GoogleDocsTable {
        private readonly List<GoogleDocsTableRow> _rows = new List<GoogleDocsTableRow>();
        private readonly List<double> _columnWidthPoints = new List<double>();

        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public string? StyleName { get; set; }
        public string? Title { get; set; }
        public string? Description { get; set; }
        public bool RepeatHeaderRow { get; set; }
        public bool HasHorizontalMerges { get; set; }
        public bool HasVerticalMerges { get; set; }
        public IReadOnlyList<GoogleDocsTableRow> Rows => _rows;
        public IReadOnlyList<double> ColumnWidthPoints => _columnWidthPoints;

        internal void AddRow(GoogleDocsTableRow row) {
            if (row == null) throw new ArgumentNullException(nameof(row));
            _rows.Add(row);
        }

        internal void AddColumnWidth(double widthPoints) {
            _columnWidthPoints.Add(widthPoints);
        }
    }

    /// <summary>
    /// Normalized table row payload used by the compiler output.
    /// </summary>
    public sealed class GoogleDocsTableRow {
        private readonly List<GoogleDocsTableCell> _cells = new List<GoogleDocsTableCell>();

        public int RowIndex { get; set; }
        public IReadOnlyList<GoogleDocsTableCell> Cells => _cells;

        internal void AddCell(GoogleDocsTableCell cell) {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            _cells.Add(cell);
        }
    }

    /// <summary>
    /// Normalized table cell payload used by the compiler output.
    /// </summary>
    public sealed class GoogleDocsTableCell {
        private readonly List<GoogleDocsParagraph> _paragraphs = new List<GoogleDocsParagraph>();

        public int ColumnIndex { get; set; }
        public int ColumnSpan { get; set; } = 1;
        public int RowSpan { get; set; } = 1;
        public string? ShadingFillColorHex { get; set; }
        public GoogleDocsTableCellBorder? LeftBorder { get; set; }
        public GoogleDocsTableCellBorder? RightBorder { get; set; }
        public GoogleDocsTableCellBorder? TopBorder { get; set; }
        public GoogleDocsTableCellBorder? BottomBorder { get; set; }
        public bool HasHorizontalMerge { get; set; }
        public bool HasVerticalMerge { get; set; }
        public IReadOnlyList<GoogleDocsParagraph> Paragraphs => _paragraphs;

        internal void AddParagraph(GoogleDocsParagraph paragraph) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            _paragraphs.Add(paragraph);
        }
    }

    public sealed class GoogleDocsTableCellBorder {
        public string? Style { get; set; }
        public string? ColorHex { get; set; }
        public uint? Size { get; set; }
    }

    public sealed class GoogleDocsParagraphBorder {
        public string? Style { get; set; }
        public string? ColorHex { get; set; }
        public uint? Size { get; set; }
        public uint? Space { get; set; }
    }

    public sealed class GoogleDocsTabStop {
        public string? Alignment { get; set; }
        public string? Leader { get; set; }
        public double OffsetPoints { get; set; }
    }

    public sealed class GoogleDocsSectionStyle {
        public string? Orientation { get; set; }
        public double? PageWidthPoints { get; set; }
        public double? PageHeightPoints { get; set; }
        public double? MarginTopPoints { get; set; }
        public double? MarginBottomPoints { get; set; }
        public double? MarginLeftPoints { get; set; }
        public double? MarginRightPoints { get; set; }
        public double? HeaderMarginPoints { get; set; }
        public double? FooterMarginPoints { get; set; }
        public int? ColumnCount { get; set; }
        public double? ColumnSpacingPoints { get; set; }
        public bool HasColumnSeparator { get; set; }
        public bool UseFirstPageHeaderFooter { get; set; }
        public int? PageNumberStart { get; set; }
    }
}
