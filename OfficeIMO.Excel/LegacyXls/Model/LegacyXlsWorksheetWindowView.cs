namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents BIFF Window2 worksheet-view metadata for one worksheet window.
    /// </summary>
    public sealed class LegacyXlsWorksheetWindowView {
        internal LegacyXlsWorksheetWindowView(
            bool showFormulas,
            bool showGridLines,
            bool showRowColumnHeadings,
            bool showZeroValues,
            bool rightToLeft,
            bool defaultGridColor,
            ushort? gridLineColorIndex,
            bool showOutlineSymbols,
            bool tabSelected,
            bool pageBreakPreview,
            bool frozenWithoutSplit,
            int? firstVisibleRow,
            int? firstVisibleColumn,
            uint? zoomScale,
            uint? zoomScaleNormal) {
            ShowFormulas = showFormulas;
            ShowGridLines = showGridLines;
            ShowRowColumnHeadings = showRowColumnHeadings;
            ShowZeroValues = showZeroValues;
            RightToLeft = rightToLeft;
            DefaultGridColor = defaultGridColor;
            GridLineColorIndex = gridLineColorIndex;
            ShowOutlineSymbols = showOutlineSymbols;
            TabSelected = tabSelected;
            PageBreakPreview = pageBreakPreview;
            FrozenWithoutSplit = frozenWithoutSplit;
            FirstVisibleRow = firstVisibleRow;
            FirstVisibleColumn = firstVisibleColumn;
            ZoomScale = zoomScale;
            ZoomScaleNormal = zoomScaleNormal;
        }

        /// <summary>
        /// Gets whether formulas are shown instead of calculated values in this worksheet window.
        /// </summary>
        public bool ShowFormulas { get; }

        /// <summary>
        /// Gets whether gridlines are visible in this worksheet window.
        /// </summary>
        public bool ShowGridLines { get; }

        /// <summary>
        /// Gets whether row and column headings are visible in this worksheet window.
        /// </summary>
        public bool ShowRowColumnHeadings { get; }

        /// <summary>
        /// Gets whether zero values are visible in this worksheet window.
        /// </summary>
        public bool ShowZeroValues { get; }

        /// <summary>
        /// Gets whether this worksheet window is displayed from right to left.
        /// </summary>
        public bool RightToLeft { get; }

        /// <summary>
        /// Gets whether this worksheet window uses the default gridline color.
        /// </summary>
        public bool DefaultGridColor { get; }

        /// <summary>
        /// Gets the BIFF color index used for gridlines when the default gridline color is disabled.
        /// </summary>
        public ushort? GridLineColorIndex { get; }

        /// <summary>
        /// Gets whether outline symbols are visible in this worksheet window.
        /// </summary>
        public bool ShowOutlineSymbols { get; }

        /// <summary>
        /// Gets whether the worksheet tab is selected in this worksheet window.
        /// </summary>
        public bool TabSelected { get; }

        /// <summary>
        /// Gets whether this worksheet window is displayed in page break preview.
        /// </summary>
        public bool PageBreakPreview { get; }

        /// <summary>
        /// Gets whether the frozen pane state uses BIFF's frozen-without-split flag.
        /// </summary>
        public bool FrozenWithoutSplit { get; }

        /// <summary>
        /// Gets the zero-based first visible row in this worksheet window.
        /// </summary>
        public int? FirstVisibleRow { get; }

        /// <summary>
        /// Gets the zero-based first visible column in this worksheet window.
        /// </summary>
        public int? FirstVisibleColumn { get; }

        /// <summary>
        /// Gets the view zoom scale percentage when it is stored by the Window2 record.
        /// </summary>
        public uint? ZoomScale { get; }

        /// <summary>
        /// Gets the normal-view zoom scale percentage stored by the Window2 record.
        /// </summary>
        public uint? ZoomScaleNormal { get; }
    }
}
