namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocParagraphFormat : IEquatable<LegacyDocParagraphFormat> {
        internal LegacyDocParagraphFormat(
            LegacyDocParagraphAlignment? alignment,
            ushort? styleIndex = null,
            int? spacingBeforeTwips = null,
            int? spacingAfterTwips = null,
            int? lineSpacingTwips = null,
            int? leftIndentTwips = null,
            int? rightIndentTwips = null,
            int? firstLineIndentTwips = null,
            bool? keepLinesTogether = null,
            bool? keepWithNext = null,
            bool? pageBreakBefore = null,
            bool? avoidWidowAndOrphan = null,
            bool? isInTable = null,
            bool? isTableTerminatingParagraph = null,
            IReadOnlyList<LegacyDocTabStop>? tabStops = null,
            IReadOnlyList<int>? tableCellWidthsTwips = null,
            int? tableLeftIndentTwips = null,
            int? tableRowHeightTwips = null,
            bool tableRowHeightIsExact = false,
            bool? tableRowCantSplit = null,
            bool? tableRowIsHeader = null,
            LegacyDocTableAlignment? tableAlignment = null,
            LegacyDocTablePreferredWidth? tablePreferredWidth = null,
            bool? tableAutofit = null,
            IReadOnlyList<LegacyDocTableCellHorizontalMerge>? tableCellHorizontalMerges = null,
            IReadOnlyList<LegacyDocTableCellVerticalMerge>? tableCellVerticalMerges = null,
            IReadOnlyList<LegacyDocTableCellVerticalAlignment>? tableCellVerticalAlignments = null,
            IReadOnlyList<bool>? tableCellFitTexts = null,
            IReadOnlyList<bool>? tableCellNoWraps = null,
            IReadOnlyList<LegacyDocTableCellMargins>? tableCellMargins = null,
            IReadOnlyList<LegacyDocTableCellShading>? tableCellShadings = null,
            IReadOnlyList<LegacyDocTableCellBorders>? tableCellBorders = null,
            LegacyDocTableCellMargins? defaultTableCellMargins = null,
            int? defaultTableCellSpacingTwips = null,
            bool hasMergedTableCells = false,
            LegacyDocParagraphShading? paragraphShading = null) {
            Alignment = alignment;
            StyleIndex = styleIndex;
            SpacingBeforeTwips = spacingBeforeTwips;
            SpacingAfterTwips = spacingAfterTwips;
            LineSpacingTwips = lineSpacingTwips;
            LeftIndentTwips = leftIndentTwips;
            RightIndentTwips = rightIndentTwips;
            FirstLineIndentTwips = firstLineIndentTwips;
            KeepLinesTogether = keepLinesTogether;
            KeepWithNext = keepWithNext;
            PageBreakBefore = pageBreakBefore;
            AvoidWidowAndOrphan = avoidWidowAndOrphan;
            IsInTable = isInTable;
            IsTableTerminatingParagraph = isTableTerminatingParagraph;
            TabStops = tabStops == null || tabStops.Count == 0
                ? Array.Empty<LegacyDocTabStop>()
                : tabStops.ToArray();
            TableCellWidthsTwips = tableCellWidthsTwips == null || tableCellWidthsTwips.Count == 0
                ? Array.Empty<int>()
                : tableCellWidthsTwips.ToArray();
            TableLeftIndentTwips = tableLeftIndentTwips.HasValue && tableLeftIndentTwips.Value > 0 && tableLeftIndentTwips.Value <= short.MaxValue
                ? tableLeftIndentTwips
                : null;
            TableRowHeightTwips = tableRowHeightTwips;
            TableRowHeightIsExact = tableRowHeightIsExact;
            TableRowCantSplit = tableRowCantSplit;
            TableRowIsHeader = tableRowIsHeader;
            TableAlignment = tableAlignment;
            TablePreferredWidth = tablePreferredWidth;
            TableAutofit = tableAutofit;
            TableCellHorizontalMerges = tableCellHorizontalMerges == null || tableCellHorizontalMerges.Count == 0
                ? Array.Empty<LegacyDocTableCellHorizontalMerge>()
                : tableCellHorizontalMerges.ToArray();
            TableCellVerticalMerges = tableCellVerticalMerges == null || tableCellVerticalMerges.Count == 0
                ? Array.Empty<LegacyDocTableCellVerticalMerge>()
                : tableCellVerticalMerges.ToArray();
            TableCellVerticalAlignments = tableCellVerticalAlignments == null || tableCellVerticalAlignments.Count == 0
                ? Array.Empty<LegacyDocTableCellVerticalAlignment>()
                : tableCellVerticalAlignments.ToArray();
            TableCellFitTexts = tableCellFitTexts == null || tableCellFitTexts.Count == 0
                ? Array.Empty<bool>()
                : tableCellFitTexts.ToArray();
            TableCellNoWraps = tableCellNoWraps == null || tableCellNoWraps.Count == 0
                ? Array.Empty<bool>()
                : tableCellNoWraps.ToArray();
            TableCellMargins = tableCellMargins == null || tableCellMargins.Count == 0
                ? Array.Empty<LegacyDocTableCellMargins>()
                : tableCellMargins.ToArray();
            TableCellShadings = tableCellShadings == null || tableCellShadings.Count == 0
                ? Array.Empty<LegacyDocTableCellShading>()
                : tableCellShadings.ToArray();
            TableCellBorders = tableCellBorders == null || tableCellBorders.Count == 0
                ? Array.Empty<LegacyDocTableCellBorders>()
                : tableCellBorders.ToArray();
            DefaultTableCellMargins = defaultTableCellMargins.HasValue && defaultTableCellMargins.Value.HasAny
                ? defaultTableCellMargins
                : null;
            DefaultTableCellSpacingTwips = defaultTableCellSpacingTwips.HasValue && defaultTableCellSpacingTwips.Value >= 0 && defaultTableCellSpacingTwips.Value <= 31680
                ? defaultTableCellSpacingTwips
                : null;
            HasMergedTableCells = hasMergedTableCells;
            ParagraphShading = paragraphShading.HasValue && paragraphShading.Value.HasAny
                ? paragraphShading
                : null;
        }

        internal LegacyDocParagraphAlignment? Alignment { get; }

        internal ushort? StyleIndex { get; }

        internal int? SpacingBeforeTwips { get; }

        internal int? SpacingAfterTwips { get; }

        internal int? LineSpacingTwips { get; }

        internal int? LeftIndentTwips { get; }

        internal int? RightIndentTwips { get; }

        internal int? FirstLineIndentTwips { get; }

        internal bool? KeepLinesTogether { get; }

        internal bool? KeepWithNext { get; }

        internal bool? PageBreakBefore { get; }

        internal bool? AvoidWidowAndOrphan { get; }

        internal bool? IsInTable { get; }

        internal bool? IsTableTerminatingParagraph { get; }

        internal IReadOnlyList<LegacyDocTabStop> TabStops { get; }

        internal IReadOnlyList<int> TableCellWidthsTwips { get; }

        internal int? TableLeftIndentTwips { get; }

        internal int? TableRowHeightTwips { get; }

        internal bool TableRowHeightIsExact { get; }

        internal bool? TableRowCantSplit { get; }

        internal bool? TableRowIsHeader { get; }

        internal LegacyDocTableAlignment? TableAlignment { get; }

        internal LegacyDocTablePreferredWidth? TablePreferredWidth { get; }

        internal bool? TableAutofit { get; }

        internal IReadOnlyList<LegacyDocTableCellHorizontalMerge> TableCellHorizontalMerges { get; }

        internal IReadOnlyList<LegacyDocTableCellVerticalMerge> TableCellVerticalMerges { get; }

        internal IReadOnlyList<LegacyDocTableCellVerticalAlignment> TableCellVerticalAlignments { get; }

        internal IReadOnlyList<bool> TableCellFitTexts { get; }

        internal IReadOnlyList<bool> TableCellNoWraps { get; }

        internal IReadOnlyList<LegacyDocTableCellMargins> TableCellMargins { get; }

        internal IReadOnlyList<LegacyDocTableCellShading> TableCellShadings { get; }

        internal IReadOnlyList<LegacyDocTableCellBorders> TableCellBorders { get; }

        internal LegacyDocTableCellMargins? DefaultTableCellMargins { get; }

        internal int? DefaultTableCellSpacingTwips { get; }

        internal bool HasMergedTableCells { get; }

        internal LegacyDocParagraphShading? ParagraphShading { get; }

        internal bool HasFormatting => Alignment != null
            || StyleIndex != null
            || SpacingBeforeTwips != null
            || SpacingAfterTwips != null
            || LineSpacingTwips != null
            || LeftIndentTwips != null
            || RightIndentTwips != null
            || FirstLineIndentTwips != null
            || KeepLinesTogether != null
            || KeepWithNext != null
            || PageBreakBefore != null
            || AvoidWidowAndOrphan != null
            || IsInTable != null
            || IsTableTerminatingParagraph != null
            || TabStops.Count > 0
            || TableCellWidthsTwips.Count > 0
            || TableLeftIndentTwips != null
            || TableRowHeightTwips != null
            || TableRowCantSplit != null
            || TableRowIsHeader != null
            || TableAlignment != null
            || TablePreferredWidth != null
            || TableAutofit != null
            || TableCellHorizontalMerges.Count > 0
            || TableCellVerticalMerges.Count > 0
            || TableCellVerticalAlignments.Count > 0
            || TableCellFitTexts.Count > 0
            || TableCellNoWraps.Count > 0
            || TableCellMargins.Count > 0
            || TableCellShadings.Count > 0
            || TableCellBorders.Count > 0
            || DefaultTableCellMargins != null
            || DefaultTableCellSpacingTwips != null
            || HasMergedTableCells
            || ParagraphShading != null;

        internal static LegacyDocParagraphFormat Default { get; } = new LegacyDocParagraphFormat(null);

        public bool Equals(LegacyDocParagraphFormat other) {
            return Alignment == other.Alignment
                && StyleIndex == other.StyleIndex
                && SpacingBeforeTwips == other.SpacingBeforeTwips
                && SpacingAfterTwips == other.SpacingAfterTwips
                && LineSpacingTwips == other.LineSpacingTwips
                && LeftIndentTwips == other.LeftIndentTwips
                && RightIndentTwips == other.RightIndentTwips
                && FirstLineIndentTwips == other.FirstLineIndentTwips
                && KeepLinesTogether == other.KeepLinesTogether
                && KeepWithNext == other.KeepWithNext
                && PageBreakBefore == other.PageBreakBefore
                && AvoidWidowAndOrphan == other.AvoidWidowAndOrphan
                && IsInTable == other.IsInTable
                && IsTableTerminatingParagraph == other.IsTableTerminatingParagraph
                && TabStopsEqual(TabStops, other.TabStops)
                && TableCellWidthsEqual(TableCellWidthsTwips, other.TableCellWidthsTwips)
                && TableLeftIndentTwips == other.TableLeftIndentTwips
                && TableRowHeightTwips == other.TableRowHeightTwips
                && TableRowHeightIsExact == other.TableRowHeightIsExact
                && TableRowCantSplit == other.TableRowCantSplit
                && TableRowIsHeader == other.TableRowIsHeader
                && TableAlignment == other.TableAlignment
                && TablePreferredWidth.Equals(other.TablePreferredWidth)
                && TableAutofit == other.TableAutofit
                && TableCellHorizontalMergesEqual(TableCellHorizontalMerges, other.TableCellHorizontalMerges)
                && TableCellVerticalMergesEqual(TableCellVerticalMerges, other.TableCellVerticalMerges)
                && TableCellVerticalAlignmentsEqual(TableCellVerticalAlignments, other.TableCellVerticalAlignments)
                && TableCellBooleansEqual(TableCellFitTexts, other.TableCellFitTexts)
                && TableCellBooleansEqual(TableCellNoWraps, other.TableCellNoWraps)
                && TableCellMarginsEqual(TableCellMargins, other.TableCellMargins)
                && TableCellShadingsEqual(TableCellShadings, other.TableCellShadings)
                && TableCellBordersEqual(TableCellBorders, other.TableCellBorders)
                && DefaultTableCellMargins.Equals(other.DefaultTableCellMargins)
                && DefaultTableCellSpacingTwips == other.DefaultTableCellSpacingTwips
                && HasMergedTableCells == other.HasMergedTableCells
                && ParagraphShading.Equals(other.ParagraphShading);
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocParagraphFormat other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Alignment.GetHashCode();
            hash = (hash * 31) + StyleIndex.GetHashCode();
            hash = (hash * 31) + SpacingBeforeTwips.GetHashCode();
            hash = (hash * 31) + SpacingAfterTwips.GetHashCode();
            hash = (hash * 31) + LineSpacingTwips.GetHashCode();
            hash = (hash * 31) + LeftIndentTwips.GetHashCode();
            hash = (hash * 31) + RightIndentTwips.GetHashCode();
            hash = (hash * 31) + FirstLineIndentTwips.GetHashCode();
            hash = (hash * 31) + KeepLinesTogether.GetHashCode();
            hash = (hash * 31) + KeepWithNext.GetHashCode();
            hash = (hash * 31) + PageBreakBefore.GetHashCode();
            hash = (hash * 31) + AvoidWidowAndOrphan.GetHashCode();
            hash = (hash * 31) + IsInTable.GetHashCode();
            hash = (hash * 31) + IsTableTerminatingParagraph.GetHashCode();
            hash = (hash * 31) + TableLeftIndentTwips.GetHashCode();
            hash = (hash * 31) + TableRowHeightTwips.GetHashCode();
            hash = (hash * 31) + TableRowHeightIsExact.GetHashCode();
            hash = (hash * 31) + TableRowCantSplit.GetHashCode();
            hash = (hash * 31) + TableRowIsHeader.GetHashCode();
            hash = (hash * 31) + TableAlignment.GetHashCode();
            hash = (hash * 31) + TablePreferredWidth.GetHashCode();
            hash = (hash * 31) + TableAutofit.GetHashCode();
            hash = (hash * 31) + DefaultTableCellSpacingTwips.GetHashCode();
            hash = (hash * 31) + HasMergedTableCells.GetHashCode();
            hash = (hash * 31) + ParagraphShading.GetHashCode();
            foreach (LegacyDocTableCellHorizontalMerge merge in TableCellHorizontalMerges) {
                hash = (hash * 31) + merge.GetHashCode();
            }

            foreach (LegacyDocTableCellVerticalMerge merge in TableCellVerticalMerges) {
                hash = (hash * 31) + merge.GetHashCode();
            }

            foreach (LegacyDocTableCellVerticalAlignment alignment in TableCellVerticalAlignments) {
                hash = (hash * 31) + alignment.GetHashCode();
            }

            foreach (bool fitText in TableCellFitTexts) {
                hash = (hash * 31) + fitText.GetHashCode();
            }

            foreach (bool noWrap in TableCellNoWraps) {
                hash = (hash * 31) + noWrap.GetHashCode();
            }

            hash = (hash * 31) + DefaultTableCellMargins.GetHashCode();
            foreach (LegacyDocTableCellMargins margins in TableCellMargins) {
                hash = (hash * 31) + margins.GetHashCode();
            }

            foreach (LegacyDocTableCellShading shading in TableCellShadings) {
                hash = (hash * 31) + shading.GetHashCode();
            }

            foreach (LegacyDocTableCellBorders borders in TableCellBorders) {
                hash = (hash * 31) + borders.GetHashCode();
            }

            foreach (LegacyDocTabStop tabStop in TabStops) {
                hash = (hash * 31) + tabStop.GetHashCode();
            }

            foreach (int width in TableCellWidthsTwips) {
                hash = (hash * 31) + width.GetHashCode();
            }

            return hash;
        }

        private static bool TableCellWidthsEqual(IReadOnlyList<int> first, IReadOnlyList<int> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (first[index] != second[index]) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellHorizontalMergesEqual(IReadOnlyList<LegacyDocTableCellHorizontalMerge> first, IReadOnlyList<LegacyDocTableCellHorizontalMerge> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (first[index] != second[index]) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellVerticalMergesEqual(IReadOnlyList<LegacyDocTableCellVerticalMerge> first, IReadOnlyList<LegacyDocTableCellVerticalMerge> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (first[index] != second[index]) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellVerticalAlignmentsEqual(IReadOnlyList<LegacyDocTableCellVerticalAlignment> first, IReadOnlyList<LegacyDocTableCellVerticalAlignment> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (first[index] != second[index]) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellBooleansEqual(IReadOnlyList<bool> first, IReadOnlyList<bool> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (first[index] != second[index]) {
                    return false;
                }
            }

            return true;
        }

        internal IReadOnlyList<LegacyDocTableCellMargins> GetTableCellMarginsForCellCount(int cellCount) {
            if (cellCount <= 0 || (TableCellMargins.Count == 0 && DefaultTableCellMargins == null)) {
                return Array.Empty<LegacyDocTableCellMargins>();
            }

            int count = Math.Max(cellCount, TableCellMargins.Count);
            var margins = new LegacyDocTableCellMargins[count];
            if (DefaultTableCellMargins != null) {
                for (int index = 0; index < count; index++) {
                    margins[index] = DefaultTableCellMargins.Value;
                }
            }

            for (int index = 0; index < TableCellMargins.Count; index++) {
                margins[index] = margins[index].Merge(TableCellMargins[index]);
            }

            return margins.Any(margin => margin.HasAny)
                ? margins
                : Array.Empty<LegacyDocTableCellMargins>();
        }

        internal IReadOnlyList<LegacyDocTableCellShading> GetTableCellShadingsForCellCount(int cellCount) {
            if (cellCount <= 0 || TableCellShadings.Count == 0) {
                return Array.Empty<LegacyDocTableCellShading>();
            }

            int count = Math.Max(cellCount, TableCellShadings.Count);
            var shadings = new LegacyDocTableCellShading[count];
            for (int index = 0; index < TableCellShadings.Count; index++) {
                shadings[index] = TableCellShadings[index];
            }

            return shadings.Any(shading => shading.HasAny)
                ? shadings
                : Array.Empty<LegacyDocTableCellShading>();
        }

        internal IReadOnlyList<LegacyDocTableCellBorders> GetTableCellBordersForCellCount(int cellCount) {
            if (cellCount <= 0 || TableCellBorders.Count == 0) {
                return Array.Empty<LegacyDocTableCellBorders>();
            }

            int count = Math.Max(cellCount, TableCellBorders.Count);
            var borders = new LegacyDocTableCellBorders[count];
            for (int index = 0; index < TableCellBorders.Count; index++) {
                borders[index] = TableCellBorders[index];
            }

            return borders.Any(border => border.HasAny)
                ? borders
                : Array.Empty<LegacyDocTableCellBorders>();
        }

        private static bool TableCellMarginsEqual(IReadOnlyList<LegacyDocTableCellMargins> first, IReadOnlyList<LegacyDocTableCellMargins> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (!first[index].Equals(second[index])) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellShadingsEqual(IReadOnlyList<LegacyDocTableCellShading> first, IReadOnlyList<LegacyDocTableCellShading> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (!first[index].Equals(second[index])) {
                    return false;
                }
            }

            return true;
        }

        private static bool TabStopsEqual(IReadOnlyList<LegacyDocTabStop> first, IReadOnlyList<LegacyDocTabStop> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (!first[index].Equals(second[index])) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellBordersEqual(IReadOnlyList<LegacyDocTableCellBorders> first, IReadOnlyList<LegacyDocTableCellBorders> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (!first[index].Equals(second[index])) {
                    return false;
                }
            }

            return true;
        }
    }
}
