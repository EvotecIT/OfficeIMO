using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// The worksheet's view: frozen panes, gridlines, zero values, headings, direction, zoom and view mode.
    /// Values are read from and written to the worksheet's first sheet view.
    /// </summary>
    public sealed class ExcelSheetView {
        private readonly ExcelSheet _sheet;

        internal ExcelSheetView(ExcelSheet sheet) {
            _sheet = sheet;
        }

        private SheetView? Element => _sheet.WorksheetRoot.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();

        /// <summary>Freezes the top rows and left columns; zero for both removes frozen panes.</summary>
        /// <param name="topRows">Number of rows at the top to freeze.</param>
        /// <param name="leftColumns">Number of columns on the left to freeze.</param>
        public ExcelSheetView Freeze(int topRows = 0, int leftColumns = 0) {
            if (topRows < 0) throw new ArgumentOutOfRangeException(nameof(topRows));
            if (leftColumns < 0) throw new ArgumentOutOfRangeException(nameof(leftColumns));
            _sheet.FreezeCore(topRows, leftColumns);
            return this;
        }

        /// <summary>Number of frozen rows; zero when no panes are frozen.</summary>
        public int FrozenRows => Inspect().FrozenRowCount;

        /// <summary>Number of frozen columns; zero when no panes are frozen.</summary>
        public int FrozenColumns => Inspect().FrozenColumnCount;

        /// <summary>Whether cell gridlines are shown.</summary>
        public bool ShowGridlines {
            get => Element?.ShowGridLines?.Value ?? true;
            set => _sheet.SetGridlinesVisibleCore(value);
        }

        /// <summary>Whether zero values are shown; false shows such cells blank.</summary>
        public bool ShowZeros {
            get => _sheet.ZeroValuesVisibleCore;
            set => _sheet.SetZeroValuesVisibleCore(value);
        }

        /// <summary>Whether row numbers and column letters are shown.</summary>
        public bool ShowRowColumnHeadings {
            get => _sheet.RowColumnHeadingsVisibleCore;
            set => _sheet.SetRowColumnHeadingsVisibleCore(value);
        }

        /// <summary>Whether the sheet is displayed from right to left.</summary>
        public bool RightToLeft {
            get => _sheet.RightToLeftCore;
            set => _sheet.SetRightToLeftCore(value);
        }

        /// <summary>Zoom percentage (10-400); null when the sheet stores none.</summary>
        public uint? ZoomScale {
            get => _sheet.GetZoomScaleCore();
            set {
                if (!value.HasValue) throw new ArgumentNullException(nameof(value), "Set a zoom percentage from 10 to 400.");
                _sheet.SetZoomScaleCore(value.Value, save: false);
            }
        }

        /// <summary>Zoom percentage (10-400) used by the normal view; null when the sheet stores none.</summary>
        public uint? ZoomScaleNormal {
            get => Element?.ZoomScaleNormal?.Value;
            set {
                if (!value.HasValue) throw new ArgumentNullException(nameof(value), "Set a zoom percentage from 10 to 400.");
                _sheet.ApplyViewOptions(new ExcelWorksheetViewOptions { ZoomScaleNormal = value });
            }
        }

        /// <summary>View mode: normal, page break preview or page layout.</summary>
        public ExcelWorksheetViewKind Kind {
            get {
                SheetViewValues? value = Element?.View?.Value;
                if (value == SheetViewValues.PageBreakPreview) return ExcelWorksheetViewKind.PageBreakPreview;
                if (value == SheetViewValues.PageLayout) return ExcelWorksheetViewKind.PageLayout;
                return ExcelWorksheetViewKind.Normal;
            }
            set => _sheet.ApplyViewOptions(new ExcelWorksheetViewOptions { View = value });
        }

        /// <summary>Reads pane, split, selection and zoom details of the sheet view.</summary>
        public ExcelWorksheetViewInfo Inspect() => _sheet.ReadViewInfo();
    }
}
