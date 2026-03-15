#pragma warning disable CS1591

namespace OfficeIMO.Word {
    public sealed class WordDocumentSnapshot {
        private readonly List<WordSectionSnapshot> _sections = new List<WordSectionSnapshot>();

        public string? FilePath { get; internal set; }
        public string? Title { get; internal set; }
        public IReadOnlyList<WordSectionSnapshot> Sections => _sections;

        internal void AddSection(WordSectionSnapshot section) {
            if (section == null) throw new ArgumentNullException(nameof(section));
            _sections.Add(section);
        }
    }

    public sealed class WordSectionSnapshot {
        private readonly List<WordBlockSnapshot> _elements = new List<WordBlockSnapshot>();

        public int Index { get; internal set; }
        public string? SectionBreakType { get; internal set; }
        public string? Orientation { get; internal set; }
        public double? PageWidthPoints { get; internal set; }
        public double? PageHeightPoints { get; internal set; }
        public double? MarginTopPoints { get; internal set; }
        public double? MarginBottomPoints { get; internal set; }
        public double? MarginLeftPoints { get; internal set; }
        public double? MarginRightPoints { get; internal set; }
        public double? HeaderMarginPoints { get; internal set; }
        public double? FooterMarginPoints { get; internal set; }
        public int? ColumnCount { get; internal set; }
        public double? ColumnSpacingPoints { get; internal set; }
        public bool HasColumnSeparator { get; internal set; }
        public int? PageNumberStart { get; internal set; }
        public int HeaderCount { get; internal set; }
        public int FooterCount { get; internal set; }
        public bool DifferentFirstPage { get; internal set; }
        public bool DifferentOddAndEvenPages { get; internal set; }
        public WordHeaderFooterSnapshot? DefaultHeader { get; internal set; }
        public WordHeaderFooterSnapshot? DefaultFooter { get; internal set; }
        public WordHeaderFooterSnapshot? FirstHeader { get; internal set; }
        public WordHeaderFooterSnapshot? FirstFooter { get; internal set; }
        public WordHeaderFooterSnapshot? EvenHeader { get; internal set; }
        public WordHeaderFooterSnapshot? EvenFooter { get; internal set; }
        public IReadOnlyList<WordBlockSnapshot> Elements => _elements;

        internal void AddElement(WordBlockSnapshot element) {
            if (element == null) throw new ArgumentNullException(nameof(element));
            _elements.Add(element);
        }
    }

    public abstract class WordBlockSnapshot {
        protected WordBlockSnapshot(string kind) {
            Kind = kind ?? throw new ArgumentNullException(nameof(kind));
        }

        public string Kind { get; }
        public int Order { get; internal set; }
    }

    public sealed class WordParagraphSnapshot : WordBlockSnapshot {
        private readonly List<WordRunSnapshot> _runs = new List<WordRunSnapshot>();
        private readonly List<WordTabStopSnapshot> _tabStops = new List<WordTabStopSnapshot>();

        public WordParagraphSnapshot() : base("paragraph") {
        }

        public string Text { get; internal set; } = string.Empty;
        public string? StyleId { get; internal set; }
        public string? StyleName { get; internal set; }
        public bool IsListItem { get; internal set; }
        public bool? IsOrderedList { get; internal set; }
        public int? ListLevel { get; internal set; }
        public string? ListStyleName { get; internal set; }
        public string? Alignment { get; internal set; }
        public double? IndentStartPoints { get; internal set; }
        public double? IndentEndPoints { get; internal set; }
        public double? IndentFirstLinePoints { get; internal set; }
        public double? SpaceAbovePoints { get; internal set; }
        public double? SpaceBelowPoints { get; internal set; }
        public int? LineSpacingValue { get; internal set; }
        public string? LineSpacingRule { get; internal set; }
        public string? ShadingFillColorHex { get; internal set; }
        public WordParagraphBorderSnapshot? LeftBorder { get; internal set; }
        public WordParagraphBorderSnapshot? RightBorder { get; internal set; }
        public WordParagraphBorderSnapshot? TopBorder { get; internal set; }
        public WordParagraphBorderSnapshot? BottomBorder { get; internal set; }
        public bool IsRightToLeft { get; internal set; }
        public bool KeepWithNext { get; internal set; }
        public bool KeepLinesTogether { get; internal set; }
        public bool AvoidWidowAndOrphan { get; internal set; }
        public bool PageBreakBefore { get; internal set; }
        public string? BookmarkName { get; internal set; }
        public int? BookmarkId { get; internal set; }
        public IReadOnlyList<WordRunSnapshot> Runs => _runs;
        public IReadOnlyList<WordTabStopSnapshot> TabStops => _tabStops;

        internal void AddRun(WordRunSnapshot run) {
            if (run == null) throw new ArgumentNullException(nameof(run));
            _runs.Add(run);
        }

        internal void AddTabStop(WordTabStopSnapshot tabStop) {
            if (tabStop == null) throw new ArgumentNullException(nameof(tabStop));
            _tabStops.Add(tabStop);
        }
    }

    public sealed class WordRunSnapshot {
        public string Text { get; internal set; } = string.Empty;
        public bool Bold { get; internal set; }
        public bool Italic { get; internal set; }
        public bool Underline { get; internal set; }
        public bool Strike { get; internal set; }
        public int? FontSize { get; internal set; }
        public string? FontFamily { get; internal set; }
        public string? ColorHex { get; internal set; }
        public string? HighlightColor { get; internal set; }
        public string? VerticalTextAlignment { get; internal set; }
        public string? CapsStyle { get; internal set; }
        public bool IsHyperlink { get; internal set; }
        public string? HyperlinkUri { get; internal set; }
        public string? HyperlinkAnchor { get; internal set; }
        public WordFootnoteSnapshot? Footnote { get; internal set; }
        public WordInlineImageSnapshot? InlineImage { get; internal set; }
    }

    public sealed class WordFootnoteSnapshot {
        private readonly List<WordParagraphSnapshot> _paragraphs = new List<WordParagraphSnapshot>();

        public long? ReferenceId { get; internal set; }
        public IReadOnlyList<WordParagraphSnapshot> Paragraphs => _paragraphs;

        internal void AddParagraph(WordParagraphSnapshot paragraph) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            _paragraphs.Add(paragraph);
        }
    }

    public sealed class WordInlineImageSnapshot {
        public string? FilePath { get; internal set; }
        public string? FileName { get; internal set; }
        public string? ContentType { get; internal set; }
        public byte[]? Bytes { get; internal set; }
        public string? Description { get; internal set; }
        public string? Title { get; internal set; }
        public double? Width { get; internal set; }
        public double? Height { get; internal set; }
        public bool IsInline { get; internal set; }
        public string? WrapText { get; internal set; }
    }

    public sealed class WordHeaderFooterSnapshot {
        private readonly List<WordBlockSnapshot> _elements = new List<WordBlockSnapshot>();

        public string Kind { get; internal set; } = string.Empty;
        public string Variant { get; internal set; } = string.Empty;
        public int TableCount { get; internal set; }
        public IReadOnlyList<WordBlockSnapshot> Elements => _elements;
        public IReadOnlyList<WordParagraphSnapshot> Paragraphs => _elements.OfType<WordParagraphSnapshot>().ToList();
        public IReadOnlyList<WordTableSnapshot> Tables => _elements.OfType<WordTableSnapshot>().ToList();

        internal void AddElement(WordBlockSnapshot element) {
            if (element == null) throw new ArgumentNullException(nameof(element));
            _elements.Add(element);
        }
    }

    public sealed class WordTableSnapshot : WordBlockSnapshot {
        private readonly List<WordTableRowSnapshot> _rows = new List<WordTableRowSnapshot>();
        private readonly List<double> _columnWidthPoints = new List<double>();

        public WordTableSnapshot() : base("table") {
        }

        public int RowCount { get; internal set; }
        public int ColumnCount { get; internal set; }
        public string? StyleName { get; internal set; }
        public string? Title { get; internal set; }
        public string? Description { get; internal set; }
        public bool RepeatHeaderRow { get; internal set; }
        public bool HasHorizontalMerges { get; internal set; }
        public bool HasVerticalMerges { get; internal set; }
        public IReadOnlyList<WordTableRowSnapshot> Rows => _rows;
        public IReadOnlyList<double> ColumnWidthPoints => _columnWidthPoints;

        internal void AddRow(WordTableRowSnapshot row) {
            if (row == null) throw new ArgumentNullException(nameof(row));
            _rows.Add(row);
        }

        internal void AddColumnWidth(double widthPoints) {
            _columnWidthPoints.Add(widthPoints);
        }
    }

    public sealed class WordTableRowSnapshot {
        private readonly List<WordTableCellSnapshot> _cells = new List<WordTableCellSnapshot>();

        public int RowIndex { get; internal set; }
        public IReadOnlyList<WordTableCellSnapshot> Cells => _cells;

        internal void AddCell(WordTableCellSnapshot cell) {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            _cells.Add(cell);
        }
    }

    public sealed class WordTableCellSnapshot {
        private readonly List<WordParagraphSnapshot> _paragraphs = new List<WordParagraphSnapshot>();

        public int ColumnIndex { get; internal set; }
        public int ColumnSpan { get; internal set; } = 1;
        public int RowSpan { get; internal set; } = 1;
        public string? ShadingFillColorHex { get; internal set; }
        public WordTableCellBorderSnapshot? LeftBorder { get; internal set; }
        public WordTableCellBorderSnapshot? RightBorder { get; internal set; }
        public WordTableCellBorderSnapshot? TopBorder { get; internal set; }
        public WordTableCellBorderSnapshot? BottomBorder { get; internal set; }
        public bool HasHorizontalMerge { get; internal set; }
        public bool HasVerticalMerge { get; internal set; }
        public IReadOnlyList<WordParagraphSnapshot> Paragraphs => _paragraphs;

        internal void AddParagraph(WordParagraphSnapshot paragraph) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            _paragraphs.Add(paragraph);
        }
    }

    public sealed class WordTableCellBorderSnapshot {
        public string? Style { get; internal set; }
        public string? ColorHex { get; internal set; }
        public uint? Size { get; internal set; }
    }

    public sealed class WordParagraphBorderSnapshot {
        public string? Style { get; internal set; }
        public string? ColorHex { get; internal set; }
        public uint? Size { get; internal set; }
        public uint? Space { get; internal set; }
    }

    public sealed class WordTabStopSnapshot {
        public string? Alignment { get; internal set; }
        public string? Leader { get; internal set; }
        public double PositionPoints { get; internal set; }
    }
}
