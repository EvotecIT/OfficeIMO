namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents workbook window geometry and display flags decoded from a legacy Window1 record.
    /// </summary>
    public sealed class LegacyXlsWorkbookWindow {
        /// <summary>
        /// Creates workbook window metadata.
        /// </summary>
        public LegacyXlsWorkbookWindow(
            short horizontalPositionTwips,
            short verticalPositionTwips,
            short widthTwips,
            short heightTwips,
            bool hidden,
            bool minimized,
            bool veryHidden,
            bool horizontalScrollBarVisible,
            bool verticalScrollBarVisible,
            bool sheetTabsVisible,
            bool autoFilterDatesGroupedChronologically,
            ushort activeSheetIndex,
            ushort firstVisibleSheetTabIndex,
            ushort selectedSheetTabCount,
            ushort sheetTabRatio) {
            HorizontalPositionTwips = horizontalPositionTwips;
            VerticalPositionTwips = verticalPositionTwips;
            WidthTwips = widthTwips;
            HeightTwips = heightTwips;
            Hidden = hidden;
            Minimized = minimized;
            VeryHidden = veryHidden;
            HorizontalScrollBarVisible = horizontalScrollBarVisible;
            VerticalScrollBarVisible = verticalScrollBarVisible;
            SheetTabsVisible = sheetTabsVisible;
            AutoFilterDatesGroupedChronologically = autoFilterDatesGroupedChronologically;
            ActiveSheetIndex = activeSheetIndex;
            FirstVisibleSheetTabIndex = firstVisibleSheetTabIndex;
            SelectedSheetTabCount = selectedSheetTabCount;
            SheetTabRatio = sheetTabRatio;
        }

        /// <summary>Gets the horizontal window position in twips.</summary>
        public short HorizontalPositionTwips { get; }

        /// <summary>Gets the vertical window position in twips.</summary>
        public short VerticalPositionTwips { get; }

        /// <summary>Gets the window width in twips.</summary>
        public short WidthTwips { get; }

        /// <summary>Gets the window height in twips.</summary>
        public short HeightTwips { get; }

        /// <summary>Gets whether the window is hidden.</summary>
        public bool Hidden { get; }

        /// <summary>Gets whether the window is minimized.</summary>
        public bool Minimized { get; }

        /// <summary>Gets whether the window is very hidden.</summary>
        public bool VeryHidden { get; }

        /// <summary>Gets whether the horizontal scroll bar is visible.</summary>
        public bool HorizontalScrollBarVisible { get; }

        /// <summary>Gets whether the vertical scroll bar is visible.</summary>
        public bool VerticalScrollBarVisible { get; }

        /// <summary>Gets whether sheet tabs are visible.</summary>
        public bool SheetTabsVisible { get; }

        /// <summary>Gets whether AutoFilter dates are grouped chronologically.</summary>
        public bool AutoFilterDatesGroupedChronologically { get; }

        /// <summary>Gets the zero-based active sheet index.</summary>
        public ushort ActiveSheetIndex { get; }

        /// <summary>Gets the zero-based first visible sheet tab index.</summary>
        public ushort FirstVisibleSheetTabIndex { get; }

        /// <summary>Gets the selected sheet tab count.</summary>
        public ushort SelectedSheetTabCount { get; }

        /// <summary>Gets the sheet tab width ratio.</summary>
        public ushort SheetTabRatio { get; }
    }
}
