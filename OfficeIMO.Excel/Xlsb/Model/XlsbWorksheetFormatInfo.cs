namespace OfficeIMO.Excel.Xlsb.Model {
    internal sealed class XlsbWorksheetFormatInfo {
        internal XlsbWorksheetFormatInfo(
            double? defaultColumnWidth,
            uint? baseColumnWidth,
            double defaultRowHeight,
            bool customDefaultRowHeight,
            bool defaultRowsHidden,
            byte maximumRowOutlineLevel,
            byte maximumColumnOutlineLevel) {
            DefaultColumnWidth = defaultColumnWidth;
            BaseColumnWidth = baseColumnWidth;
            DefaultRowHeight = defaultRowHeight;
            CustomDefaultRowHeight = customDefaultRowHeight;
            DefaultRowsHidden = defaultRowsHidden;
            MaximumRowOutlineLevel = maximumRowOutlineLevel;
            MaximumColumnOutlineLevel = maximumColumnOutlineLevel;
        }

        internal double? DefaultColumnWidth { get; }

        internal uint? BaseColumnWidth { get; }

        internal double DefaultRowHeight { get; }

        internal bool CustomDefaultRowHeight { get; }

        internal bool DefaultRowsHidden { get; }

        internal byte MaximumRowOutlineLevel { get; }

        internal byte MaximumColumnOutlineLevel { get; }
    }
}
