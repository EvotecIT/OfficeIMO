namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private ExcelSheetPageSetup? _pageSetup;
        private ExcelSheetView? _view;

        /// <summary>Paper, orientation, scaling, margins, print area and titles, print options, headers and footers.</summary>
        public ExcelSheetPageSetup PageSetup => _pageSetup ??= new ExcelSheetPageSetup(this);

        /// <summary>Frozen panes, gridlines, zero values, headings, direction, zoom and view mode.</summary>
        public ExcelSheetView View => _view ??= new ExcelSheetView(this);
    }
}
