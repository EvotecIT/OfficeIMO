namespace OfficeIMO.Excel.Xlsb.Model {
    internal sealed class XlsbPaneInfo {
        internal XlsbPaneInfo(
            double horizontalSplit,
            double verticalSplit,
            int topRow,
            int leftColumn,
            uint activePane,
            bool frozen,
            bool frozenNoSplit) {
            HorizontalSplit = horizontalSplit;
            VerticalSplit = verticalSplit;
            TopRow = topRow;
            LeftColumn = leftColumn;
            ActivePane = activePane;
            Frozen = frozen;
            FrozenNoSplit = frozenNoSplit;
        }

        internal double HorizontalSplit { get; }

        internal double VerticalSplit { get; }

        internal int TopRow { get; }

        internal int LeftColumn { get; }

        internal uint ActivePane { get; }

        internal bool Frozen { get; }

        internal bool FrozenNoSplit { get; }
    }
}
