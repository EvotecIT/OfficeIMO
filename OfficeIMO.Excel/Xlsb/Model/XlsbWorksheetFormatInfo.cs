namespace OfficeIMO.Excel.Xlsb.Model {
    internal sealed class XlsbWorksheetFormatInfo {
        internal XlsbWorksheetFormatInfo(
            double defaultColumnWidth,
            double defaultRowHeight,
            bool customDefaultRowHeight,
            bool defaultRowsHidden,
            byte maximumRowOutlineLevel,
            byte maximumColumnOutlineLevel) {
            DefaultColumnWidth = defaultColumnWidth;
            DefaultRowHeight = defaultRowHeight;
            CustomDefaultRowHeight = customDefaultRowHeight;
            DefaultRowsHidden = defaultRowsHidden;
            MaximumRowOutlineLevel = maximumRowOutlineLevel;
            MaximumColumnOutlineLevel = maximumColumnOutlineLevel;
        }

        internal double DefaultColumnWidth { get; }

        internal double DefaultRowHeight { get; }

        internal bool CustomDefaultRowHeight { get; }

        internal bool DefaultRowsHidden { get; }

        internal byte MaximumRowOutlineLevel { get; }

        internal byte MaximumColumnOutlineLevel { get; }
    }
}
