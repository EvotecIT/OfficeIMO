namespace OfficeIMO.Excel.Xlsb.Model {
    internal sealed class XlsbColumnInfo {
        internal XlsbColumnInfo(
            int firstColumn,
            int lastColumn,
            double width,
            uint styleIndex,
            bool hidden,
            bool userSet,
            bool bestFit,
            bool phonetic,
            byte outlineLevel,
            bool collapsed) {
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
            Width = width;
            StyleIndex = styleIndex;
            Hidden = hidden;
            UserSet = userSet;
            BestFit = bestFit;
            Phonetic = phonetic;
            OutlineLevel = outlineLevel;
            Collapsed = collapsed;
        }

        internal int FirstColumn { get; }

        internal int LastColumn { get; }

        internal double Width { get; }

        internal uint StyleIndex { get; }

        internal bool Hidden { get; }

        internal bool UserSet { get; }

        internal bool BestFit { get; }

        internal bool Phonetic { get; }

        internal byte OutlineLevel { get; }

        internal bool Collapsed { get; }
    }
}
