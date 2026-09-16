namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Base request emitted by the Google Docs batch compiler.
    /// </summary>
    public abstract class GoogleDocsRequest {
        /// <summary>Initializes the provider-neutral request kind.</summary>
        protected GoogleDocsRequest(string kind) {
            Kind = kind ?? throw new ArgumentNullException(nameof(kind));
        }

        /// <summary>Gets the compiled request kind.</summary>
        public string Kind { get; }
        /// <summary>Gets or sets the zero-based source section index.</summary>
        public int SectionIndex { get; set; }
        /// <summary>Gets or sets the source element index within its section.</summary>
        public int ElementIndex { get; set; }
        /// <summary>Gets or sets section formatting associated with the request.</summary>
        public GoogleDocsSectionStyle? SectionStyle { get; set; }
    }

    /// <summary>
    /// Inserts one paragraph block into the target Google Doc.
    /// </summary>
    public sealed class GoogleDocsInsertParagraphRequest : GoogleDocsRequest {
        /// <summary>Creates an insert-paragraph request.</summary>
        public GoogleDocsInsertParagraphRequest() : base("insertParagraph") {
        }

        /// <summary>Gets or sets the paragraph payload to insert.</summary>
        public GoogleDocsParagraph Paragraph { get; set; } = new GoogleDocsParagraph();
    }

    /// <summary>
    /// Inserts one table block into the target Google Doc.
    /// </summary>
    public sealed class GoogleDocsInsertTableRequest : GoogleDocsRequest {
        /// <summary>Creates an insert-table request.</summary>
        public GoogleDocsInsertTableRequest() : base("insertTable") {
        }

        /// <summary>Gets or sets the table payload to insert.</summary>
        public GoogleDocsTable Table { get; set; } = new GoogleDocsTable();
        /// <summary>Gets or sets whether a section break precedes this table.</summary>
        public bool StartsNewSectionBefore { get; set; }
        /// <summary>Gets or sets the section-break type when a break precedes this table.</summary>
        public string? SectionBreakType { get; set; }
    }

    /// <summary>
    /// Normalized header/footer segment payload compiled from OfficeIMO snapshot data.
    /// </summary>
    public sealed class GoogleDocsSegment {
        private readonly List<GoogleDocsRequest> _requests = new List<GoogleDocsRequest>();

        /// <summary>Gets or sets the zero-based source section index.</summary>
        public int SectionIndex { get; set; }
        /// <summary>Gets or sets the segment kind, such as header or footer.</summary>
        public string Kind { get; set; } = string.Empty;
        /// <summary>Gets or sets the header/footer variant.</summary>
        public string Variant { get; set; } = string.Empty;
        /// <summary>Gets or sets the number of tables compiled into this segment.</summary>
        public int TableCount { get; set; }
        /// <summary>Gets the compiled requests in segment order.</summary>
        public IReadOnlyList<GoogleDocsRequest> Requests => _requests;
        /// <summary>Projects the paragraph payloads from this segment's requests.</summary>
        public IReadOnlyList<GoogleDocsParagraph> Paragraphs => _requests
            .OfType<GoogleDocsInsertParagraphRequest>()
            .Select(request => request.Paragraph)
            .ToList();
        /// <summary>Projects the table payloads from this segment's requests.</summary>
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

        /// <summary>Gets or sets the paragraph text.</summary>
        public string Text { get; set; } = string.Empty;
        /// <summary>Gets or sets the source paragraph-style identifier.</summary>
        public string? StyleId { get; set; }
        /// <summary>Gets or sets the source paragraph-style name.</summary>
        public string? StyleName { get; set; }
        /// <summary>Gets or sets whether a section break precedes this paragraph.</summary>
        public bool StartsNewSectionBefore { get; set; }
        /// <summary>Gets or sets the preceding section-break type.</summary>
        public string? SectionBreakType { get; set; }
        /// <summary>Gets or sets whether this paragraph is a list item.</summary>
        public bool IsListItem { get; set; }
        /// <summary>Gets or sets ordered-list status when this is a list item.</summary>
        public bool? IsOrderedList { get; set; }
        /// <summary>Gets or sets the list nesting level.</summary>
        public int? ListLevel { get; set; }
        /// <summary>Gets or sets the source list-style name.</summary>
        public string? ListStyleName { get; set; }
        /// <summary>Gets or sets paragraph alignment.</summary>
        public string? Alignment { get; set; }
        /// <summary>Gets or sets the start indent in points.</summary>
        public double? IndentStartPoints { get; set; }
        /// <summary>Gets or sets the end indent in points.</summary>
        public double? IndentEndPoints { get; set; }
        /// <summary>Gets or sets the first-line indent in points.</summary>
        public double? IndentFirstLinePoints { get; set; }
        /// <summary>Gets or sets space above in points.</summary>
        public double? SpaceAbovePoints { get; set; }
        /// <summary>Gets or sets space below in points.</summary>
        public double? SpaceBelowPoints { get; set; }
        /// <summary>Gets or sets line spacing as a percentage.</summary>
        public double? LineSpacingPercent { get; set; }
        /// <summary>Gets or sets paragraph shading as a hex color.</summary>
        public string? ShadingFillColorHex { get; set; }
        /// <summary>Gets or sets the left paragraph border.</summary>
        public GoogleDocsParagraphBorder? LeftBorder { get; set; }
        /// <summary>Gets or sets the right paragraph border.</summary>
        public GoogleDocsParagraphBorder? RightBorder { get; set; }
        /// <summary>Gets or sets the top paragraph border.</summary>
        public GoogleDocsParagraphBorder? TopBorder { get; set; }
        /// <summary>Gets or sets the bottom paragraph border.</summary>
        public GoogleDocsParagraphBorder? BottomBorder { get; set; }
        /// <summary>Gets or sets right-to-left paragraph direction.</summary>
        public bool IsRightToLeft { get; set; }
        /// <summary>Gets or sets whether this paragraph stays with the next.</summary>
        public bool KeepWithNext { get; set; }
        /// <summary>Gets or sets whether paragraph lines stay together.</summary>
        public bool KeepLinesTogether { get; set; }
        /// <summary>Gets or sets widow/orphan avoidance.</summary>
        public bool AvoidWidowAndOrphan { get; set; }
        /// <summary>Gets or sets whether a page break precedes the paragraph.</summary>
        public bool PageBreakBefore { get; set; }
        /// <summary>Gets or sets an associated bookmark name.</summary>
        public string? BookmarkName { get; set; }
        /// <summary>Gets or sets an associated bookmark identifier.</summary>
        public int? BookmarkId { get; set; }
        /// <summary>Gets the compiled text and inline-object runs.</summary>
        public IReadOnlyList<GoogleDocsParagraphRun> Runs => _runs;
        /// <summary>Gets the compiled tab stops.</summary>
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
        /// <summary>Gets or sets the run text.</summary>
        public string Text { get; set; } = string.Empty;
        /// <summary>Gets or sets bold formatting.</summary>
        public bool Bold { get; set; }
        /// <summary>Gets or sets italic formatting.</summary>
        public bool Italic { get; set; }
        /// <summary>Gets or sets underline formatting.</summary>
        public bool Underline { get; set; }
        /// <summary>Gets or sets strikethrough formatting.</summary>
        public bool Strike { get; set; }
        /// <summary>Gets or sets the font size.</summary>
        public int? FontSize { get; set; }
        /// <summary>Gets or sets the font family.</summary>
        public string? FontFamily { get; set; }
        /// <summary>Gets or sets foreground color as hex text.</summary>
        public string? ForegroundColorHex { get; set; }
        /// <summary>Gets or sets a named highlight color; unrecognized names are ignored by the exporter.</summary>
        public string? HighlightColor { get; set; }
        /// <summary>Gets or sets superscript, subscript, or baseline alignment.</summary>
        public string? VerticalTextAlignment { get; set; }
        /// <summary>Gets or sets the source capitalization style.</summary>
        public string? CapsStyle { get; set; }
        /// <summary>Gets or sets a link attached to this run.</summary>
        public GoogleDocsLink? Link { get; set; }
        /// <summary>Gets or sets a footnote attached to this run.</summary>
        public GoogleDocsFootnote? Footnote { get; set; }
        /// <summary>Gets or sets an inline image attached to this run.</summary>
        public GoogleDocsInlineImage? InlineImage { get; set; }
    }

    /// <summary>
    /// Normalized hyperlink payload used by the compiler output.
    /// </summary>
    public sealed class GoogleDocsLink {
        /// <summary>Gets or sets an external link URI.</summary>
        public string? Uri { get; set; }
        /// <summary>Gets or sets an internal bookmark or anchor target.</summary>
        public string? Anchor { get; set; }
    }

    /// <summary>
    /// Normalized inline image payload used by the compiler output.
    /// </summary>
    public sealed class GoogleDocsInlineImage {
        /// <summary>Gets or sets the source image file path, when available.</summary>
        public string? FilePath { get; set; }
        /// <summary>Gets or sets the source file name, when available.</summary>
        public string? FileName { get; set; }
        /// <summary>Gets or sets the image media type, when available.</summary>
        public string? ContentType { get; set; }
        /// <summary>Gets or sets image bytes; the payload does not copy the array.</summary>
        public byte[]? Bytes { get; set; }
        /// <summary>Gets or sets alternative description text.</summary>
        public string? Description { get; set; }
        /// <summary>Gets or sets an optional image title.</summary>
        public string? Title { get; set; }
        /// <summary>Gets or sets the requested width.</summary>
        public double? Width { get; set; }
        /// <summary>Gets or sets the requested height.</summary>
        public double? Height { get; set; }
        /// <summary>Gets or sets whether the source image is inline.</summary>
        public bool IsInline { get; set; }
        /// <summary>Gets or sets the source text-wrapping mode, when available.</summary>
        public string? WrapText { get; set; }
    }

    /// <summary>
    /// Normalized footnote payload used by the compiler output.
    /// </summary>
    public sealed class GoogleDocsFootnote {
        private readonly List<GoogleDocsParagraph> _paragraphs = new List<GoogleDocsParagraph>();

        /// <summary>Gets or sets the source footnote reference identifier.</summary>
        public long? ReferenceId { get; set; }
        /// <summary>Gets the footnote's compiled paragraphs.</summary>
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

        /// <summary>Gets or sets the number of table rows.</summary>
        public int RowCount { get; set; }
        /// <summary>Gets or sets the number of table columns.</summary>
        public int ColumnCount { get; set; }
        /// <summary>Gets or sets the source table-style name.</summary>
        public string? StyleName { get; set; }
        /// <summary>Gets or sets an optional table title.</summary>
        public string? Title { get; set; }
        /// <summary>Gets or sets an optional table description.</summary>
        public string? Description { get; set; }
        /// <summary>Gets or sets whether the header row repeats across pages.</summary>
        public bool RepeatHeaderRow { get; set; }
        /// <summary>Gets or sets whether the source table contains horizontal merges.</summary>
        public bool HasHorizontalMerges { get; set; }
        /// <summary>Gets or sets whether the source table contains vertical merges.</summary>
        public bool HasVerticalMerges { get; set; }
        /// <summary>Gets compiled table rows.</summary>
        public IReadOnlyList<GoogleDocsTableRow> Rows => _rows;
        /// <summary>Gets source column widths in points.</summary>
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

        /// <summary>Gets or sets the zero-based source row index.</summary>
        public int RowIndex { get; set; }
        /// <summary>Gets compiled cells in this row.</summary>
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

        /// <summary>Gets or sets the zero-based source column index.</summary>
        public int ColumnIndex { get; set; }
        /// <summary>Gets or sets the number of columns spanned; defaults to one.</summary>
        public int ColumnSpan { get; set; } = 1;
        /// <summary>Gets or sets the number of rows spanned; defaults to one.</summary>
        public int RowSpan { get; set; } = 1;
        /// <summary>Gets or sets cell shading as a hex color.</summary>
        public string? ShadingFillColorHex { get; set; }
        /// <summary>Gets or sets the left cell border.</summary>
        public GoogleDocsTableCellBorder? LeftBorder { get; set; }
        /// <summary>Gets or sets the right cell border.</summary>
        public GoogleDocsTableCellBorder? RightBorder { get; set; }
        /// <summary>Gets or sets the top cell border.</summary>
        public GoogleDocsTableCellBorder? TopBorder { get; set; }
        /// <summary>Gets or sets the bottom cell border.</summary>
        public GoogleDocsTableCellBorder? BottomBorder { get; set; }
        /// <summary>Gets or sets whether the cell continues a horizontal merge.</summary>
        public bool HasHorizontalMerge { get; set; }
        /// <summary>Gets or sets whether the cell continues a vertical merge.</summary>
        public bool HasVerticalMerge { get; set; }
        /// <summary>Gets compiled paragraphs within the cell.</summary>
        public IReadOnlyList<GoogleDocsParagraph> Paragraphs => _paragraphs;

        internal void AddParagraph(GoogleDocsParagraph paragraph) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            _paragraphs.Add(paragraph);
        }
    }

    /// <summary>Normalized border on one table-cell side.</summary>
    public sealed class GoogleDocsTableCellBorder {
        /// <summary>Gets or sets the border style.</summary>
        public string? Style { get; set; }
        /// <summary>Gets or sets the border color as hex text.</summary>
        public string? ColorHex { get; set; }
        /// <summary>Gets or sets the source border size.</summary>
        public uint? Size { get; set; }
    }

    /// <summary>Normalized border on one paragraph side.</summary>
    public sealed class GoogleDocsParagraphBorder {
        /// <summary>Gets or sets the border style.</summary>
        public string? Style { get; set; }
        /// <summary>Gets or sets the border color as hex text.</summary>
        public string? ColorHex { get; set; }
        /// <summary>Gets or sets the source border size.</summary>
        public uint? Size { get; set; }
        /// <summary>Gets or sets the source border spacing.</summary>
        public uint? Space { get; set; }
    }

    /// <summary>A tab stop compiled from paragraph formatting.</summary>
    public sealed class GoogleDocsTabStop {
        /// <summary>Gets or sets tab alignment.</summary>
        public string? Alignment { get; set; }
        /// <summary>Gets or sets the tab leader style.</summary>
        public string? Leader { get; set; }
        /// <summary>Gets or sets the tab offset in points.</summary>
        public double OffsetPoints { get; set; }
    }

    /// <summary>Page and section layout projected from a Word section.</summary>
    public sealed class GoogleDocsSectionStyle {
        /// <summary>Gets or sets the requested page orientation.</summary>
        public string? Orientation { get; set; }
        /// <summary>Gets or sets page width in points.</summary>
        public double? PageWidthPoints { get; set; }
        /// <summary>Gets or sets page height in points.</summary>
        public double? PageHeightPoints { get; set; }
        /// <summary>Gets or sets top margin in points.</summary>
        public double? MarginTopPoints { get; set; }
        /// <summary>Gets or sets bottom margin in points.</summary>
        public double? MarginBottomPoints { get; set; }
        /// <summary>Gets or sets left margin in points.</summary>
        public double? MarginLeftPoints { get; set; }
        /// <summary>Gets or sets right margin in points.</summary>
        public double? MarginRightPoints { get; set; }
        /// <summary>Gets or sets header offset in points.</summary>
        public double? HeaderMarginPoints { get; set; }
        /// <summary>Gets or sets footer offset in points.</summary>
        public double? FooterMarginPoints { get; set; }
        /// <summary>Gets or sets the source section's column count.</summary>
        public int? ColumnCount { get; set; }
        /// <summary>Gets or sets spacing between source columns in points.</summary>
        public double? ColumnSpacingPoints { get; set; }
        /// <summary>Gets or sets whether the source columns have a separator.</summary>
        public bool HasColumnSeparator { get; set; }
        /// <summary>Gets or sets whether first-page header/footer variants are used.</summary>
        public bool UseFirstPageHeaderFooter { get; set; }
        /// <summary>Gets or sets the starting page number, when specified.</summary>
        public int? PageNumberStart { get; set; }
    }
}
