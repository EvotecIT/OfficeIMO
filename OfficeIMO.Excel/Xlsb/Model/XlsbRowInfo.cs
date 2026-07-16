namespace OfficeIMO.Excel.Xlsb.Model {
    internal sealed class XlsbRowInfo {
        private readonly List<XlsbColumnSpan> _spans = new List<XlsbColumnSpan>();

        internal XlsbRowInfo(
            int row,
            uint styleIndex,
            ushort heightTwips,
            byte outlineLevel,
            bool collapsed,
            bool hidden,
            bool customHeight,
            bool customFormat,
            bool phonetic) {
            Row = row;
            StyleIndex = styleIndex;
            HeightTwips = heightTwips;
            OutlineLevel = outlineLevel;
            Collapsed = collapsed;
            Hidden = hidden;
            CustomHeight = customHeight;
            CustomFormat = customFormat;
            Phonetic = phonetic;
        }

        internal int Row { get; }

        internal uint StyleIndex { get; }

        internal ushort HeightTwips { get; }

        internal byte OutlineLevel { get; }

        internal bool Collapsed { get; }

        internal bool Hidden { get; }

        internal bool CustomHeight { get; }

        internal bool CustomFormat { get; }

        internal bool Phonetic { get; }

        internal IReadOnlyList<XlsbColumnSpan> Spans => _spans;

        internal void AddSpan(int firstColumn, int lastColumn) {
            _spans.Add(new XlsbColumnSpan(firstColumn, lastColumn));
        }

        internal bool ContainsZeroBasedColumn(int column) {
            return _spans.Any(span => span.FirstColumn <= column && column <= span.LastColumn);
        }
    }

    internal readonly struct XlsbColumnSpan {
        internal XlsbColumnSpan(int firstColumn, int lastColumn) {
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
        }

        internal int FirstColumn { get; }

        internal int LastColumn { get; }
    }
}
