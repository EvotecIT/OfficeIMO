namespace OfficeIMO.Word.LegacyDoc.Model {
    internal enum LegacyDocTableCellHorizontalMerge {
        None,
        Restart,
        Continue
    }

    internal enum LegacyDocTableCellVerticalMerge {
        None,
        Restart,
        Continue
    }

    internal enum LegacyDocTableCellVerticalAlignment {
        Top,
        Center,
        Bottom
    }

    internal enum LegacyDocTableCellTextDirection {
        LeftToRightTopToBottom,
        TopToBottomRightToLeft,
        BottomToTopLeftToRight,
        LeftToRightTopToBottomRotated,
        TopToBottomRightToLeftRotated
    }

    internal enum LegacyDocTableAlignment {
        Left,
        Center,
        Right
    }

    internal enum LegacyDocTablePreferredWidthUnit {
        Auto,
        Percent,
        Dxa
    }

    internal readonly struct LegacyDocTablePreferredWidth : IEquatable<LegacyDocTablePreferredWidth> {
        internal LegacyDocTablePreferredWidth(LegacyDocTablePreferredWidthUnit unit, int value) {
            Unit = unit;
            Value = value;
        }

        internal LegacyDocTablePreferredWidthUnit Unit { get; }

        internal int Value { get; }

        public bool Equals(LegacyDocTablePreferredWidth other) {
            return Unit == other.Unit && Value == other.Value;
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocTablePreferredWidth other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Unit.GetHashCode();
            hash = (hash * 31) + Value.GetHashCode();
            return hash;
        }
    }

    internal readonly struct LegacyDocTableCellShading : IEquatable<LegacyDocTableCellShading> {
        internal LegacyDocTableCellShading(string? fillColorHex) {
            FillColorHex = string.IsNullOrWhiteSpace(fillColorHex)
                ? null
                : fillColorHex!.Replace("#", string.Empty).ToLowerInvariant();
        }

        internal string? FillColorHex { get; }

        internal bool HasAny => !string.IsNullOrEmpty(FillColorHex);

        public bool Equals(LegacyDocTableCellShading other) {
            return string.Equals(FillColorHex, other.FillColorHex, StringComparison.Ordinal);
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocTableCellShading other && Equals(other);
        }

        public override int GetHashCode() {
            return FillColorHex == null ? 0 : FillColorHex.GetHashCode();
        }
    }

    internal enum LegacyDocTableCellBorderStyle {
        None,
        Single,
        Double,
        Dotted,
        Dashed
    }

    internal readonly struct LegacyDocTableCellBorder : IEquatable<LegacyDocTableCellBorder> {
        internal LegacyDocTableCellBorder(LegacyDocTableCellBorderStyle style, string? colorHex, int sizeEighthPoints, int spacePoints) {
            Style = style;
            ColorHex = string.IsNullOrWhiteSpace(colorHex)
                ? null
                : colorHex!.Replace("#", string.Empty).ToLowerInvariant();
            SizeEighthPoints = sizeEighthPoints;
            SpacePoints = spacePoints;
        }

        internal LegacyDocTableCellBorderStyle Style { get; }

        internal string? ColorHex { get; }

        internal int SizeEighthPoints { get; }

        internal int SpacePoints { get; }

        internal bool HasAny => Style != LegacyDocTableCellBorderStyle.None;

        public bool Equals(LegacyDocTableCellBorder other) {
            return Style == other.Style
                && string.Equals(ColorHex, other.ColorHex, StringComparison.Ordinal)
                && SizeEighthPoints == other.SizeEighthPoints
                && SpacePoints == other.SpacePoints;
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocTableCellBorder other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Style.GetHashCode();
            hash = (hash * 31) + (ColorHex == null ? 0 : ColorHex.GetHashCode());
            hash = (hash * 31) + SizeEighthPoints.GetHashCode();
            hash = (hash * 31) + SpacePoints.GetHashCode();
            return hash;
        }
    }

    internal readonly struct LegacyDocTableCellBorders : IEquatable<LegacyDocTableCellBorders> {
        internal LegacyDocTableCellBorders(
            LegacyDocTableCellBorder top,
            LegacyDocTableCellBorder left,
            LegacyDocTableCellBorder bottom,
            LegacyDocTableCellBorder right) {
            Top = top;
            Left = left;
            Bottom = bottom;
            Right = right;
        }

        internal LegacyDocTableCellBorder Top { get; }

        internal LegacyDocTableCellBorder Left { get; }

        internal LegacyDocTableCellBorder Bottom { get; }

        internal LegacyDocTableCellBorder Right { get; }

        internal bool HasAny => Top.HasAny || Left.HasAny || Bottom.HasAny || Right.HasAny;

        public bool Equals(LegacyDocTableCellBorders other) {
            return Top.Equals(other.Top)
                && Left.Equals(other.Left)
                && Bottom.Equals(other.Bottom)
                && Right.Equals(other.Right);
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocTableCellBorders other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Top.GetHashCode();
            hash = (hash * 31) + Left.GetHashCode();
            hash = (hash * 31) + Bottom.GetHashCode();
            hash = (hash * 31) + Right.GetHashCode();
            return hash;
        }
    }

    internal readonly struct LegacyDocTableCellMargins : IEquatable<LegacyDocTableCellMargins> {
        internal LegacyDocTableCellMargins(int? topTwips, int? rightTwips, int? bottomTwips, int? leftTwips) {
            TopTwips = topTwips;
            RightTwips = rightTwips;
            BottomTwips = bottomTwips;
            LeftTwips = leftTwips;
        }

        internal int? TopTwips { get; }

        internal int? RightTwips { get; }

        internal int? BottomTwips { get; }

        internal int? LeftTwips { get; }

        internal bool HasAny => TopTwips != null || RightTwips != null || BottomTwips != null || LeftTwips != null;

        internal LegacyDocTableCellMargins Merge(LegacyDocTableCellMargins margins) {
            return new LegacyDocTableCellMargins(
                margins.TopTwips ?? TopTwips,
                margins.RightTwips ?? RightTwips,
                margins.BottomTwips ?? BottomTwips,
                margins.LeftTwips ?? LeftTwips);
        }

        public bool Equals(LegacyDocTableCellMargins other) {
            return TopTwips == other.TopTwips
                && RightTwips == other.RightTwips
                && BottomTwips == other.BottomTwips
                && LeftTwips == other.LeftTwips;
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocTableCellMargins other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + TopTwips.GetHashCode();
            hash = (hash * 31) + RightTwips.GetHashCode();
            hash = (hash * 31) + BottomTwips.GetHashCode();
            hash = (hash * 31) + LeftTwips.GetHashCode();
            return hash;
        }
    }

    internal abstract class LegacyDocBodyBlock {
        private protected LegacyDocBodyBlock() {
        }
    }

    internal sealed class LegacyDocParagraphBlock : LegacyDocBodyBlock {
        internal LegacyDocParagraphBlock(
            IReadOnlyList<LegacyDocTextRun> runs,
            LegacyDocParagraphFormat format,
            int startCharacter,
            int endCharacter,
            IReadOnlyList<LegacyDocBookmark>? bookmarks = null) {
            Runs = runs;
            Format = format;
            StartCharacter = startCharacter;
            EndCharacter = endCharacter;
            Bookmarks = bookmarks ?? Array.Empty<LegacyDocBookmark>();
        }

        internal IReadOnlyList<LegacyDocTextRun> Runs { get; }

        internal LegacyDocParagraphFormat Format { get; }

        internal int StartCharacter { get; }

        internal int EndCharacter { get; }

        internal IReadOnlyList<LegacyDocBookmark> Bookmarks { get; }
    }

    internal sealed class LegacyDocSectionBreakBlock : LegacyDocBodyBlock {
        internal LegacyDocSectionBreakBlock(LegacyDocSectionFormat format) {
            Format = format;
        }

        internal LegacyDocSectionFormat Format { get; }
    }

    internal sealed class LegacyDocTableBlock : LegacyDocBodyBlock {
        internal LegacyDocTableBlock(
            IReadOnlyList<LegacyDocTableRow> rows,
            int startCharacter,
            int endCharacter,
            IReadOnlyList<LegacyDocBookmark>? bookmarks = null) {
            Rows = rows;
            StartCharacter = startCharacter;
            EndCharacter = endCharacter;
            Bookmarks = bookmarks ?? Array.Empty<LegacyDocBookmark>();
        }

        internal IReadOnlyList<LegacyDocTableRow> Rows { get; }

        internal int StartCharacter { get; }

        internal int EndCharacter { get; }

        internal IReadOnlyList<LegacyDocBookmark> Bookmarks { get; }
    }

    internal sealed class LegacyDocTableRow {
        internal LegacyDocTableRow(
            IReadOnlyList<LegacyDocTableCell> cells,
            IReadOnlyList<int>? cellWidthsTwips = null,
            int? tableLeftIndentTwips = null,
            int? rowHeightTwips = null,
            bool rowHeightIsExact = false,
            bool? rowCantSplit = null,
            bool? rowIsHeader = null,
            LegacyDocTableAlignment? tableAlignment = null,
            IReadOnlyList<LegacyDocTableCellHorizontalMerge>? cellHorizontalMerges = null,
            IReadOnlyList<LegacyDocTableCellVerticalMerge>? cellVerticalMerges = null,
            IReadOnlyList<LegacyDocTableCellVerticalAlignment>? cellVerticalAlignments = null,
            IReadOnlyList<LegacyDocTableCellTextDirection>? cellTextDirections = null,
            IReadOnlyList<bool>? cellFitTexts = null,
            IReadOnlyList<bool>? cellNoWraps = null,
            IReadOnlyList<bool>? cellHideMarks = null,
            IReadOnlyList<LegacyDocTableCellMargins>? cellMargins = null,
            IReadOnlyList<LegacyDocTableCellShading>? cellShadings = null,
            IReadOnlyList<LegacyDocTableCellBorders>? cellBorders = null,
            int? defaultCellSpacingTwips = null,
            LegacyDocTablePreferredWidth? tablePreferredWidth = null,
            bool? tableAutofit = null,
            IReadOnlyList<LegacyDocBookmark>? bookmarksBefore = null) {
            Cells = cells;
            BookmarksBefore = bookmarksBefore == null || bookmarksBefore.Count == 0
                ? Array.Empty<LegacyDocBookmark>()
                : bookmarksBefore.ToArray();
            CellWidthsTwips = cellWidthsTwips == null || cellWidthsTwips.Count == 0
                ? Array.Empty<int>()
                : cellWidthsTwips.ToArray();
            TableLeftIndentTwips = tableLeftIndentTwips.HasValue && tableLeftIndentTwips.Value > 0 && tableLeftIndentTwips.Value <= short.MaxValue
                ? tableLeftIndentTwips
                : null;
            RowHeightTwips = rowHeightTwips;
            RowHeightIsExact = rowHeightIsExact;
            RowCantSplit = rowCantSplit;
            RowIsHeader = rowIsHeader;
            TableAlignment = tableAlignment;
            CellHorizontalMerges = cellHorizontalMerges == null || cellHorizontalMerges.Count == 0
                ? Array.Empty<LegacyDocTableCellHorizontalMerge>()
                : cellHorizontalMerges.ToArray();
            CellVerticalMerges = cellVerticalMerges == null || cellVerticalMerges.Count == 0
                ? Array.Empty<LegacyDocTableCellVerticalMerge>()
                : cellVerticalMerges.ToArray();
            CellVerticalAlignments = cellVerticalAlignments == null || cellVerticalAlignments.Count == 0
                ? Array.Empty<LegacyDocTableCellVerticalAlignment>()
                : cellVerticalAlignments.ToArray();
            CellTextDirections = cellTextDirections == null || cellTextDirections.Count == 0
                ? Array.Empty<LegacyDocTableCellTextDirection>()
                : cellTextDirections.ToArray();
            CellFitTexts = cellFitTexts == null || cellFitTexts.Count == 0
                ? Array.Empty<bool>()
                : cellFitTexts.ToArray();
            CellNoWraps = cellNoWraps == null || cellNoWraps.Count == 0
                ? Array.Empty<bool>()
                : cellNoWraps.ToArray();
            CellHideMarks = cellHideMarks == null || cellHideMarks.Count == 0
                ? Array.Empty<bool>()
                : cellHideMarks.ToArray();
            CellMargins = cellMargins == null || cellMargins.Count == 0
                ? Array.Empty<LegacyDocTableCellMargins>()
                : cellMargins.ToArray();
            CellShadings = cellShadings == null || cellShadings.Count == 0
                ? Array.Empty<LegacyDocTableCellShading>()
                : cellShadings.ToArray();
            CellBorders = cellBorders == null || cellBorders.Count == 0
                ? Array.Empty<LegacyDocTableCellBorders>()
                : cellBorders.ToArray();
            DefaultCellSpacingTwips = defaultCellSpacingTwips.HasValue && defaultCellSpacingTwips.Value >= 0 && defaultCellSpacingTwips.Value <= 31680
                ? defaultCellSpacingTwips
                : null;
            TablePreferredWidth = tablePreferredWidth;
            TableAutofit = tableAutofit;
        }

        internal IReadOnlyList<LegacyDocTableCell> Cells { get; }

        internal IReadOnlyList<LegacyDocBookmark> BookmarksBefore { get; }

        internal IReadOnlyList<int> CellWidthsTwips { get; }

        internal int? TableLeftIndentTwips { get; }

        internal int? RowHeightTwips { get; }

        internal bool RowHeightIsExact { get; }

        internal bool? RowCantSplit { get; }

        internal bool? RowIsHeader { get; }

        internal LegacyDocTableAlignment? TableAlignment { get; }

        internal IReadOnlyList<LegacyDocTableCellHorizontalMerge> CellHorizontalMerges { get; }

        internal IReadOnlyList<LegacyDocTableCellVerticalMerge> CellVerticalMerges { get; }

        internal IReadOnlyList<LegacyDocTableCellVerticalAlignment> CellVerticalAlignments { get; }

        internal IReadOnlyList<LegacyDocTableCellTextDirection> CellTextDirections { get; }

        internal IReadOnlyList<bool> CellFitTexts { get; }

        internal IReadOnlyList<bool> CellNoWraps { get; }

        internal IReadOnlyList<bool> CellHideMarks { get; }

        internal IReadOnlyList<LegacyDocTableCellMargins> CellMargins { get; }

        internal IReadOnlyList<LegacyDocTableCellShading> CellShadings { get; }

        internal IReadOnlyList<LegacyDocTableCellBorders> CellBorders { get; }

        internal int? DefaultCellSpacingTwips { get; }

        internal LegacyDocTablePreferredWidth? TablePreferredWidth { get; }

        internal bool? TableAutofit { get; }
    }

    internal sealed class LegacyDocTableCell {
        internal LegacyDocTableCell(IReadOnlyList<LegacyDocTextRun> runs, LegacyDocParagraphFormat format) {
            Paragraphs = new[] { new LegacyDocTableCellParagraph(runs, format) };
            Runs = runs;
            Format = format;
            Text = string.Concat(Paragraphs.Select(paragraph => paragraph.Text));
        }

        internal LegacyDocTableCell(IReadOnlyList<LegacyDocTableCellParagraph> paragraphs) {
            Paragraphs = paragraphs.Count == 0
                ? new[] { new LegacyDocTableCellParagraph(Array.Empty<LegacyDocTextRun>(), LegacyDocParagraphFormat.Default) }
                : paragraphs.ToArray();
            Runs = Paragraphs[0].Runs;
            Format = Paragraphs[0].Format;
            Text = string.Concat(Paragraphs.Select(paragraph => paragraph.Text));
        }

        internal string Text { get; }

        internal IReadOnlyList<LegacyDocTableCellParagraph> Paragraphs { get; }

        internal IReadOnlyList<LegacyDocTextRun> Runs { get; }

        internal LegacyDocParagraphFormat Format { get; }
    }

    internal sealed class LegacyDocTableCellParagraph {
        internal LegacyDocTableCellParagraph(IReadOnlyList<LegacyDocTextRun> runs, LegacyDocParagraphFormat format)
            : this(runs, format, GetStartCharacter(runs), GetEndCharacter(runs)) {
        }

        internal LegacyDocTableCellParagraph(
            IReadOnlyList<LegacyDocTextRun> runs,
            LegacyDocParagraphFormat format,
            int startCharacter,
            int endCharacter,
            IReadOnlyList<LegacyDocBookmark>? bookmarks = null) {
            Runs = runs;
            Format = format;
            StartCharacter = startCharacter;
            EndCharacter = endCharacter;
            Bookmarks = bookmarks ?? Array.Empty<LegacyDocBookmark>();
            Text = string.Concat(runs.Select(run => run.Text));
        }

        internal string Text { get; }

        internal IReadOnlyList<LegacyDocTextRun> Runs { get; }

        internal LegacyDocParagraphFormat Format { get; }

        internal int StartCharacter { get; }

        internal int EndCharacter { get; }

        internal IReadOnlyList<LegacyDocBookmark> Bookmarks { get; }

        private static int GetStartCharacter(IReadOnlyList<LegacyDocTextRun> runs) {
            foreach (LegacyDocTextRun run in runs) {
                if (run.CharacterPositions.Count > 0) {
                    return run.CharacterPositions[0];
                }
            }

            return 0;
        }

        private static int GetEndCharacter(IReadOnlyList<LegacyDocTextRun> runs) {
            for (int index = runs.Count - 1; index >= 0; index--) {
                IReadOnlyList<int> positions = runs[index].CharacterPositions;
                if (positions.Count > 0) {
                    return positions[positions.Count - 1] + 1;
                }
            }

            return 0;
        }
    }
}
