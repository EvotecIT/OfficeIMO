using System;

namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Convenience helpers for creating a navigable index (Table of Contents) sheet
    /// and wiring back-links across the workbook. Thin wrapper over ExcelDocument APIs.
    /// </summary>
    public static class SheetIndex {
        /// <summary>
        /// Adds a styled Table of Contents sheet with hyperlinks to each sheet.
        /// </summary>
        /// <param name="workbook">Target workbook.</param>
        /// <param name="sheetName">Name of the TOC sheet. Default: "TOC".</param>
        /// <param name="placeFirst">Move TOC to the first position. Default: true.</param>
        /// <param name="includeNamedRanges">Whether to add a column listing named ranges per sheet.</param>
        /// <param name="includeHiddenNamedRanges">Include hidden named ranges in the listing.</param>
        public static void Add(ExcelDocument workbook, string sheetName = "TOC", bool placeFirst = true,
            bool includeNamedRanges = false, bool includeHiddenNamedRanges = false)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            workbook.AddTableOfContents(sheetName: sheetName, placeFirst: placeFirst, withHyperlinks: true,
                includeNamedRanges: includeNamedRanges, includeHiddenNamedRanges: includeHiddenNamedRanges,
                rangeNameFilter: null, styled: true);
        }

        /// <summary>
        /// Adds a small back-link (e.g., "← TOC") at the given cell on every non-TOC sheet.
        /// </summary>
        public static void AddBackLinks(ExcelDocument workbook, string tocSheetName = "TOC", int row = 2, int col = 1, string text = "\u2190 TOC")
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            workbook.AddBackLinksToToc(tocSheetName, row, col, text);
        }
    }
}

