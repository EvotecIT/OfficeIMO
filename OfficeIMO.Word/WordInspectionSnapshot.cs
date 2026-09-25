
namespace OfficeIMO.Word {
    /// <summary>Detached structural and metadata snapshot of a Word document.</summary>
    public sealed class WordDocumentSnapshot {
        private readonly List<WordSectionSnapshot> _sections = new List<WordSectionSnapshot>();

        /// <summary>Gets the file path currently associated with the document, when one exists.</summary>
        public string? FilePath { get; internal set; }
        /// <summary>Gets the document title from core properties.</summary>
        public string? Title { get; internal set; }
        /// <summary>Gets the document author from core properties.</summary>
        public string? Author { get; internal set; }
        /// <summary>Gets the document subject from core properties.</summary>
        public string? Subject { get; internal set; }
        /// <summary>Gets document keywords from core properties.</summary>
        public string? Keywords { get; internal set; }
        /// <summary>Gets sections in document order.</summary>
        public IReadOnlyList<WordSectionSnapshot> Sections => _sections;

        internal void AddSection(WordSectionSnapshot section) {
            if (section == null) throw new ArgumentNullException(nameof(section));
            _sections.Add(section);
        }
    }

    /// <summary>Page setup, headers, footers, and body elements for one Word section.</summary>
    public sealed class WordSectionSnapshot {
        private readonly List<WordBlockSnapshot> _elements = new List<WordBlockSnapshot>();

        /// <summary>Gets the zero-based section index.</summary>
        public int Index { get; internal set; }
        /// <summary>Gets the resolved section-break type, including Word's default when no type was authored.</summary>
        public string? SectionBreakType { get; internal set; }
        /// <summary>Gets page orientation.</summary>
        public string? Orientation { get; internal set; }
        /// <summary>Gets page width in points.</summary>
        public double? PageWidthPoints { get; internal set; }
        /// <summary>Gets page height in points.</summary>
        public double? PageHeightPoints { get; internal set; }
        /// <summary>Gets the top page margin in points.</summary>
        public double? MarginTopPoints { get; internal set; }
        /// <summary>Gets the bottom page margin in points.</summary>
        public double? MarginBottomPoints { get; internal set; }
        /// <summary>Gets the left page margin in points.</summary>
        public double? MarginLeftPoints { get; internal set; }
        /// <summary>Gets the right page margin in points.</summary>
        public double? MarginRightPoints { get; internal set; }
        /// <summary>Gets the header distance from the page edge in points.</summary>
        public double? HeaderMarginPoints { get; internal set; }
        /// <summary>Gets the footer distance from the page edge in points.</summary>
        public double? FooterMarginPoints { get; internal set; }
        /// <summary>Gets the section column count.</summary>
        public int? ColumnCount { get; internal set; }
        /// <summary>Gets spacing between columns in points.</summary>
        public double? ColumnSpacingPoints { get; internal set; }
        /// <summary>Gets whether a rule separates section columns.</summary>
        public bool HasColumnSeparator { get; internal set; }
        /// <summary>Gets the explicit starting page number.</summary>
        public int? PageNumberStart { get; internal set; }
        /// <summary>Gets the number of header parts referenced by the section.</summary>
        public int HeaderCount { get; internal set; }
        /// <summary>Gets the number of footer parts referenced by the section.</summary>
        public int FooterCount { get; internal set; }
        /// <summary>Gets whether the first page uses distinct headers and footers.</summary>
        public bool DifferentFirstPage { get; internal set; }
        /// <summary>Gets whether the section uses an explicit or inherited even-page header or footer while the document-wide odd/even setting is enabled.</summary>
        public bool DifferentOddAndEvenPages { get; internal set; }
        /// <summary>Gets the default header snapshot.</summary>
        public WordHeaderFooterSnapshot? DefaultHeader { get; internal set; }
        /// <summary>Gets the default footer snapshot.</summary>
        public WordHeaderFooterSnapshot? DefaultFooter { get; internal set; }
        /// <summary>Gets the first-page header snapshot.</summary>
        public WordHeaderFooterSnapshot? FirstHeader { get; internal set; }
        /// <summary>Gets the first-page footer snapshot.</summary>
        public WordHeaderFooterSnapshot? FirstFooter { get; internal set; }
        /// <summary>Gets the even-page header snapshot.</summary>
        public WordHeaderFooterSnapshot? EvenHeader { get; internal set; }
        /// <summary>Gets the even-page footer snapshot.</summary>
        public WordHeaderFooterSnapshot? EvenFooter { get; internal set; }
        /// <summary>Gets body paragraphs and tables in document order.</summary>
        public IReadOnlyList<WordBlockSnapshot> Elements => _elements;

        internal void AddElement(WordBlockSnapshot element) {
            if (element == null) throw new ArgumentNullException(nameof(element));
            _elements.Add(element);
        }
    }

    /// <summary>Base snapshot for an ordered Word body element.</summary>
    public abstract class WordBlockSnapshot {
        /// <summary>Initializes a block with its stable kind name.</summary>
        protected WordBlockSnapshot(string kind) {
            Kind = kind ?? throw new ArgumentNullException(nameof(kind));
        }

        /// <summary>Gets the stable block kind, such as <c>paragraph</c> or <c>table</c>.</summary>
        public string Kind { get; }
        /// <summary>Gets the zero-based order within the containing story.</summary>
        public int Order { get; internal set; }
    }

    /// <summary>Text, formatting, list, border, bookmark, and pagination state for one paragraph.</summary>
    public sealed class WordParagraphSnapshot : WordBlockSnapshot {
        private readonly List<WordRunSnapshot> _runs = new List<WordRunSnapshot>();
        private readonly List<WordInlineFieldSnapshot> _fields = new List<WordInlineFieldSnapshot>();
        private readonly List<WordTabStopSnapshot> _tabStops = new List<WordTabStopSnapshot>();

        /// <summary>Creates an empty paragraph snapshot.</summary>
        public WordParagraphSnapshot() : base("paragraph") {
        }

        /// <summary>Gets text extracted from the paragraph's runs, including text hidden by run formatting.</summary>
        public string Text { get; internal set; } = string.Empty;
        /// <summary>Gets the applied paragraph style identifier.</summary>
        public string? StyleId { get; internal set; }
        /// <summary>Gets the OfficeIMO built-in paragraph-style classification, or the document style name for a custom classification when available; otherwise <c>Custom</c>.</summary>
        public string? StyleName { get; internal set; }
        /// <summary>Gets whether the paragraph participates in numbering.</summary>
        public bool IsListItem { get; internal set; }
        /// <summary>Gets whether list numbering is ordered, or <see langword="null"/> when not resolved.</summary>
        public bool? IsOrderedList { get; internal set; }
        /// <summary>Gets the zero-based list nesting level.</summary>
        public int? ListLevel { get; internal set; }
        /// <summary>Gets the OfficeIMO list-style classification.</summary>
        public string? ListStyleName { get; internal set; }
        /// <summary>Gets the paragraph alignment authored directly on the paragraph, without resolving inherited styles.</summary>
        public string? Alignment { get; internal set; }
        /// <summary>Gets leading indentation in points.</summary>
        public double? IndentStartPoints { get; internal set; }
        /// <summary>Gets trailing indentation in points.</summary>
        public double? IndentEndPoints { get; internal set; }
        /// <summary>Gets first-line indentation in points; negative values represent hanging indents.</summary>
        public double? IndentFirstLinePoints { get; internal set; }
        /// <summary>Gets spacing above the paragraph in points.</summary>
        public double? SpaceAbovePoints { get; internal set; }
        /// <summary>Gets spacing below the paragraph in points.</summary>
        public double? SpaceBelowPoints { get; internal set; }
        /// <summary>Gets the raw Word line-spacing value.</summary>
        public int? LineSpacingValue { get; internal set; }
        /// <summary>Gets the authored line-spacing rule.</summary>
        public string? LineSpacingRule { get; internal set; }
        /// <summary>Gets the authored OOXML paragraph-shading fill token, typically an RGB hex value but possibly the automatic-color keyword.</summary>
        public string? ShadingFillColorHex { get; internal set; }
        /// <summary>Paragraph shading pattern, when explicitly authored.</summary>
        public WordShadingPattern? ShadingPattern { get; internal set; }
        /// <summary>Gets the left paragraph border.</summary>
        public WordParagraphBorderSnapshot? LeftBorder { get; internal set; }
        /// <summary>Gets the right paragraph border.</summary>
        public WordParagraphBorderSnapshot? RightBorder { get; internal set; }
        /// <summary>Gets the top paragraph border.</summary>
        public WordParagraphBorderSnapshot? TopBorder { get; internal set; }
        /// <summary>Gets the bottom paragraph border.</summary>
        public WordParagraphBorderSnapshot? BottomBorder { get; internal set; }
        /// <summary>Gets directly authored right-to-left layout without resolving paragraph styles.</summary>
        public bool IsRightToLeft { get; internal set; }
        /// <summary>Gets directly authored keep-with-next formatting without resolving paragraph styles.</summary>
        public bool KeepWithNext { get; internal set; }
        /// <summary>Gets directly authored keep-lines-together formatting without resolving paragraph styles.</summary>
        public bool KeepLinesTogether { get; internal set; }
        /// <summary>Gets directly authored widow and orphan control without resolving paragraph styles.</summary>
        public bool AvoidWidowAndOrphan { get; internal set; }
        /// <summary>Gets directly authored page-break-before formatting without resolving paragraph styles.</summary>
        public bool PageBreakBefore { get; internal set; }
        /// <summary>Gets the first bookmark name associated with the paragraph.</summary>
        public string? BookmarkName { get; internal set; }
        /// <summary>Gets the first bookmark identifier associated with the paragraph.</summary>
        public int? BookmarkId { get; internal set; }
        /// <summary>Gets text runs in source order.</summary>
        public IReadOnlyList<WordRunSnapshot> Runs => _runs;
        /// <summary>Simple fields and their positions among the paragraph's ordinary runs.</summary>
        public IReadOnlyList<WordInlineFieldSnapshot> InlineFields => _fields;
        /// <summary>Gets explicit paragraph tab stops.</summary>
        public IReadOnlyList<WordTabStopSnapshot> TabStops => _tabStops;

        internal void AddRun(WordRunSnapshot run) {
            if (run == null) throw new ArgumentNullException(nameof(run));
            _runs.Add(run);
        }

        internal void AddInlineField(WordInlineFieldSnapshot field) {
            if (field == null) throw new ArgumentNullException(nameof(field));
            _fields.Add(field);
        }

        internal void AddTabStop(WordTabStopSnapshot tabStop) {
            if (tabStop == null) throw new ArgumentNullException(nameof(tabStop));
            _tabStops.Add(tabStop);
        }
    }

    /// <summary>A simple Word field in a paragraph's inline order.</summary>
    public sealed class WordInlineFieldSnapshot {
        /// <summary>Number of ordinary runs before this field.</summary>
        public int RunIndex { get; internal set; }
        /// <summary>Raw field instruction.</summary>
        public string Instruction { get; internal set; } = string.Empty;
        /// <summary>Cached displayed result.</summary>
        public string ResultText { get; internal set; } = string.Empty;
        /// <summary>Whether the field is locked.</summary>
        public bool IsLocked { get; internal set; }
        /// <summary>Whether the cached value is marked for refresh.</summary>
        public bool IsDirty { get; internal set; }
        /// <summary>Whether the result contains content beyond plain text runs.</summary>
        public bool HasUnsupportedResultContent { get; internal set; }
        /// <summary>Whether the field is nested in markup that cannot be represented as a native ODT field.</summary>
        public bool HasUnsupportedContainer { get; internal set; }
    }

    /// <summary>Extracted text, character formatting, links, notes, and images for one Word run.</summary>
    public sealed class WordRunSnapshot {
        internal IReadOnlyDictionary<int, WordBreakType>? NonTextBreaks { get; set; }
        internal IReadOnlyList<WordPositionedImageSnapshot> PositionedImages { get; set; } = Array.Empty<WordPositionedImageSnapshot>();
        /// <summary>Gets text extracted from the run, including text hidden by run formatting.</summary>
        public string Text { get; internal set; } = string.Empty;
        /// <summary>Gets whether bold formatting is authored directly on the run.</summary>
        public bool Bold { get; internal set; }
        /// <summary>Gets whether italic formatting is authored directly on the run.</summary>
        public bool Italic { get; internal set; }
        /// <summary>Gets whether underline formatting is authored directly on the run.</summary>
        public bool Underline { get; internal set; }
        /// <summary>Gets the authored underline style.</summary>
        public WordUnderlineStyle? UnderlineStyle { get; internal set; }
        /// <summary>Gets whether single or double strikethrough is authored directly on the run.</summary>
        public bool Strike { get; internal set; }
        /// <summary>Gets whether double strikethrough is authored directly on the run.</summary>
        public bool DoubleStrike { get; internal set; }
        /// <summary>Gets the directly authored integral compatibility font size in points, truncated from half-point precision.</summary>
        public int? FontSize { get; internal set; }
        /// <summary>Gets the directly authored run font size in points with Word's native half-point precision.</summary>
        public double? FontSizePoints { get; internal set; }
        /// <summary>Gets the font family authored directly on the run, without resolving inherited styles.</summary>
        public string? FontFamily { get; internal set; }
        /// <summary>Gets the OOXML text-color token authored directly on the run.</summary>
        public string? ColorHex { get; internal set; }
        /// <summary>Gets the named Word highlight color.</summary>
        public string? HighlightColor { get; internal set; }
        /// <summary>Gets the OOXML run-shading fill token authored directly on the run.</summary>
        public string? RunShadingFillColorHex { get; internal set; }
        /// <summary>Run shading pattern, when explicitly authored.</summary>
        public WordShadingPattern? RunShadingPattern { get; internal set; }
        /// <summary>Gets vertical alignment such as superscript or subscript.</summary>
        public string? VerticalTextAlignment { get; internal set; }
        /// <summary>Gets capitalization formatting such as caps or small caps.</summary>
        public string? CapsStyle { get; internal set; }
        /// <summary>Gets whether the run belongs to a hyperlink.</summary>
        public bool IsHyperlink { get; internal set; }
        /// <summary>Gets the external hyperlink URI.</summary>
        public string? HyperlinkUri { get; internal set; }
        /// <summary>Gets the internal bookmark target.</summary>
        public string? HyperlinkAnchor { get; internal set; }
        /// <summary>Gets the referenced footnote content.</summary>
        public WordFootnoteSnapshot? Footnote { get; internal set; }
        /// <summary>Gets the referenced endnote content.</summary>
        public WordEndnoteSnapshot? Endnote { get; internal set; }
        /// <summary>Gets the first inline or positioned image encountered in the run.</summary>
        public WordInlineImageSnapshot? InlineImage { get; internal set; }
    }

    internal sealed class WordPositionedImageSnapshot {
        internal WordPositionedImageSnapshot(int offset, WordInlineImageSnapshot image) {
            Offset = offset;
            Image = image;
        }
        internal int Offset { get; }
        internal WordInlineImageSnapshot Image { get; }
    }

    /// <summary>Paragraph content for one referenced Word footnote.</summary>
    public sealed class WordFootnoteSnapshot {
        private readonly List<WordParagraphSnapshot> _paragraphs = new List<WordParagraphSnapshot>();

        /// <summary>Gets the footnote identifier from the reference element.</summary>
        public long? ReferenceId { get; internal set; }
        /// <summary>Gets footnote paragraphs in source order.</summary>
        public IReadOnlyList<WordParagraphSnapshot> Paragraphs => _paragraphs;

        internal void AddParagraph(WordParagraphSnapshot paragraph) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            _paragraphs.Add(paragraph);
        }
    }

    /// <summary>Paragraph content for one referenced Word endnote.</summary>
    public sealed class WordEndnoteSnapshot {
        private readonly List<WordParagraphSnapshot> _paragraphs = new List<WordParagraphSnapshot>();

        /// <summary>Gets the endnote identifier from the reference element.</summary>
        public long? ReferenceId { get; internal set; }
        /// <summary>Gets endnote paragraphs in source order.</summary>
        public IReadOnlyList<WordParagraphSnapshot> Paragraphs => _paragraphs;

        internal void AddParagraph(WordParagraphSnapshot paragraph) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            _paragraphs.Add(paragraph);
        }
    }

    /// <summary>Embedded image content, dimensions, accessibility text, and layout.</summary>
    public sealed class WordInlineImageSnapshot {
        /// <summary>Gets available image-origin text, such as a caller-supplied path or file name, a linked URI, or no value for package-loaded images.</summary>
        public string? FilePath { get; internal set; }
        /// <summary>Gets the image file name.</summary>
        public string? FileName { get; internal set; }
        /// <summary>Gets the image MIME type inferred from the available file name or path extension, or <see langword="null"/> when it cannot be inferred.</summary>
        public string? ContentType { get; internal set; }
        /// <summary>Gets the embedded image bytes.</summary>
        public byte[]? Bytes { get; internal set; }
        /// <summary>Gets image alternative text.</summary>
        public string? Description { get; internal set; }
        /// <summary>Gets the image title.</summary>
        public string? Title { get; internal set; }
        /// <summary>Gets rendered width in pixels at 96 pixels per inch.</summary>
        public double? Width { get; internal set; }
        /// <summary>Gets rendered height in pixels at 96 pixels per inch.</summary>
        public double? Height { get; internal set; }
        /// <summary>Gets whether the image is inline with text instead of anchored.</summary>
        public bool IsInline { get; internal set; }
        /// <summary>Gets the text-wrapping mode for an anchored image.</summary>
        public string? WrapText { get; internal set; }
    }

    /// <summary>Ordered paragraphs and tables from one header or footer part.</summary>
    public sealed class WordHeaderFooterSnapshot {
        private readonly List<WordBlockSnapshot> _elements = new List<WordBlockSnapshot>();

        /// <summary>Gets whether the story is a header or footer.</summary>
        public string Kind { get; internal set; } = string.Empty;
        /// <summary>Gets the default, first-page, or even-page variant.</summary>
        public string Variant { get; internal set; } = string.Empty;
        /// <summary>Gets the number of tables in the story.</summary>
        public int TableCount { get; internal set; }
        /// <summary>Gets paragraphs and tables in source order.</summary>
        public IReadOnlyList<WordBlockSnapshot> Elements => _elements;
        /// <summary>Gets paragraph snapshots in source order.</summary>
        public IReadOnlyList<WordParagraphSnapshot> Paragraphs => _elements.OfType<WordParagraphSnapshot>().ToList();
        /// <summary>Gets table snapshots in source order.</summary>
        public IReadOnlyList<WordTableSnapshot> Tables => _elements.OfType<WordTableSnapshot>().ToList();

        internal void AddElement(WordBlockSnapshot element) {
            if (element == null) throw new ArgumentNullException(nameof(element));
            _elements.Add(element);
        }
    }

    /// <summary>Structure, accessibility metadata, merges, and widths for one Word table.</summary>
    public sealed class WordTableSnapshot : WordBlockSnapshot {
        private readonly List<WordTableRowSnapshot> _rows = new List<WordTableRowSnapshot>();
        private readonly List<double> _columnWidthPoints = new List<double>();

        /// <summary>Creates an empty table snapshot.</summary>
        public WordTableSnapshot() : base("table") {
        }

        /// <summary>Gets the row count.</summary>
        public int RowCount { get; internal set; }
        /// <summary>Gets the maximum number of physical cell elements in any row; inspect each cell's span for logical layout.</summary>
        public int ColumnCount { get; internal set; }
        /// <summary>Gets the OfficeIMO built-in table-style classification; snapshot creation requires an authored table style identifier to be recognized by the built-in mapping.</summary>
        public string? StyleName { get; internal set; }
        /// <summary>Gets the table title used by assistive technology.</summary>
        public string? Title { get; internal set; }
        /// <summary>Gets the table description used by assistive technology.</summary>
        public string? Description { get; internal set; }
        /// <summary>Gets whether a header row is configured to repeat on subsequent pages.</summary>
        public bool RepeatHeaderRow { get; internal set; }
        /// <summary>Gets whether at least one cell uses explicit horizontal-merge markup; grid-span-only cells do not set this value.</summary>
        public bool HasHorizontalMerges { get; internal set; }
        /// <summary>Gets whether at least one cell spans multiple rows.</summary>
        public bool HasVerticalMerges { get; internal set; }
        /// <summary>Gets rows in source order.</summary>
        public IReadOnlyList<WordTableRowSnapshot> Rows => _rows;
        /// <summary>Gets authored grid-column widths in points.</summary>
        public IReadOnlyList<double> ColumnWidthPoints => _columnWidthPoints;

        internal void AddRow(WordTableRowSnapshot row) {
            if (row == null) throw new ArgumentNullException(nameof(row));
            _rows.Add(row);
        }

        internal void AddColumnWidth(double widthPoints) {
            _columnWidthPoints.Add(widthPoints);
        }
    }

    /// <summary>Physical cell elements from one table row, in source order.</summary>
    public sealed class WordTableRowSnapshot {
        private readonly List<WordTableCellSnapshot> _cells = new List<WordTableCellSnapshot>();

        /// <summary>Gets the zero-based row index.</summary>
        public int RowIndex { get; internal set; }
        /// <summary>Gets physical cell elements in source order.</summary>
        public IReadOnlyList<WordTableCellSnapshot> Cells => _cells;

        internal void AddCell(WordTableCellSnapshot cell) {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            _cells.Add(cell);
        }
    }

    /// <summary>Logical span, borders, shading, and paragraph content for one table cell.</summary>
    public sealed class WordTableCellSnapshot {
        private readonly List<WordParagraphSnapshot> _paragraphs = new List<WordParagraphSnapshot>();

        /// <summary>Gets the zero-based physical cell index within the row.</summary>
        public int ColumnIndex { get; internal set; }
        /// <summary>Gets the number of logical columns occupied by the cell.</summary>
        public int ColumnSpan { get; internal set; } = 1;
        /// <summary>Gets the number of logical rows occupied by the cell.</summary>
        public int RowSpan { get; internal set; } = 1;
        /// <summary>Gets the authored OOXML cell-shading fill token, typically an RGB hex value but possibly the automatic-color keyword.</summary>
        public string? ShadingFillColorHex { get; internal set; }
        /// <summary>Gets the left cell border.</summary>
        public WordTableCellBorderSnapshot? LeftBorder { get; internal set; }
        /// <summary>Gets the right cell border.</summary>
        public WordTableCellBorderSnapshot? RightBorder { get; internal set; }
        /// <summary>Gets the top cell border.</summary>
        public WordTableCellBorderSnapshot? TopBorder { get; internal set; }
        /// <summary>Gets the bottom cell border.</summary>
        public WordTableCellBorderSnapshot? BottomBorder { get; internal set; }
        /// <summary>Gets whether the cell participates in explicit horizontal-merge markup; a grid-span-only cell does not set this value.</summary>
        public bool HasHorizontalMerge { get; internal set; }
        /// <summary>Gets whether the cell participates in a vertical merge.</summary>
        public bool HasVerticalMerge { get; internal set; }
        /// <summary>Gets cell paragraphs in source order.</summary>
        public IReadOnlyList<WordParagraphSnapshot> Paragraphs => _paragraphs;

        internal void AddParagraph(WordParagraphSnapshot paragraph) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            _paragraphs.Add(paragraph);
        }
    }

    /// <summary>Style, color, and width of one table-cell border.</summary>
    public sealed class WordTableCellBorderSnapshot {
        /// <summary>Gets the authored Word border style.</summary>
        public string? Style { get; internal set; }
        /// <summary>Gets the authored OOXML border-color token, typically an RGB hex value but possibly the automatic-color keyword.</summary>
        public string? ColorHex { get; internal set; }
        /// <summary>Gets the border width in eighths of a point.</summary>
        public uint? Size { get; internal set; }
    }

    /// <summary>Style, color, width, and spacing of one paragraph border.</summary>
    public sealed class WordParagraphBorderSnapshot {
        /// <summary>Gets the authored Word border style.</summary>
        public string? Style { get; internal set; }
        /// <summary>Gets the authored OOXML border-color token, typically an RGB hex value but possibly the automatic-color keyword.</summary>
        public string? ColorHex { get; internal set; }
        /// <summary>Gets the border width in eighths of a point.</summary>
        public uint? Size { get; internal set; }
        /// <summary>Gets spacing between the border and paragraph in points.</summary>
        public uint? Space { get; internal set; }
    }

    /// <summary>Alignment, leader, and position of an explicit paragraph tab stop.</summary>
    public sealed class WordTabStopSnapshot {
        /// <summary>Gets tab-stop alignment.</summary>
        public string? Alignment { get; internal set; }
        /// <summary>Gets the leader character style.</summary>
        public string? Leader { get; internal set; }
        /// <summary>Gets tab position from the leading margin in points.</summary>
        public double PositionPoints { get; internal set; }
    }
}
