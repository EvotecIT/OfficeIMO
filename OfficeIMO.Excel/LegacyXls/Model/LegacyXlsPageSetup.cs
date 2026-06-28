namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents worksheet page setup metadata parsed from legacy XLS print records.
    /// </summary>
    public sealed class LegacyXlsPageSetup {
        /// <summary>Gets the left margin in inches, when present.</summary>
        public double? LeftMargin { get; internal set; }

        /// <summary>Gets the right margin in inches, when present.</summary>
        public double? RightMargin { get; internal set; }

        /// <summary>Gets the top margin in inches, when present.</summary>
        public double? TopMargin { get; internal set; }

        /// <summary>Gets the bottom margin in inches, when present.</summary>
        public double? BottomMargin { get; internal set; }

        /// <summary>Gets the header margin in inches, when present.</summary>
        public double? HeaderMargin { get; internal set; }

        /// <summary>Gets the footer margin in inches, when present.</summary>
        public double? FooterMargin { get; internal set; }

        /// <summary>Gets the raw Excel header/footer control string for the worksheet header, when present.</summary>
        public string? HeaderText { get; internal set; }

        /// <summary>Gets the raw Excel header/footer control string for the worksheet footer, when present.</summary>
        public string? FooterText { get; internal set; }

        /// <summary>Gets the raw Excel header/footer control string for the first-page header, when present.</summary>
        public string? FirstHeaderText { get; internal set; }

        /// <summary>Gets the raw Excel header/footer control string for the first-page footer, when present.</summary>
        public string? FirstFooterText { get; internal set; }

        /// <summary>Gets the raw Excel header/footer control string for even-page headers, when present.</summary>
        public string? EvenHeaderText { get; internal set; }

        /// <summary>Gets the raw Excel header/footer control string for even-page footers, when present.</summary>
        public string? EvenFooterText { get; internal set; }

        /// <summary>Gets whether odd and even pages use different headers and footers, when present.</summary>
        public bool? DifferentOddEvenHeaderFooter { get; internal set; }

        /// <summary>Gets whether the first page uses different headers and footers, when present.</summary>
        public bool? DifferentFirstHeaderFooter { get; internal set; }

        /// <summary>Gets whether headers and footers scale with the document, when present.</summary>
        public bool? ScaleHeaderFooterWithDocument { get; internal set; }

        /// <summary>Gets whether headers and footers align with page margins, when present.</summary>
        public bool? AlignHeaderFooterWithMargins { get; internal set; }

        /// <summary>Gets whether worksheet gridlines should print, when present.</summary>
        public bool? PrintGridLines { get; internal set; }

        /// <summary>Gets whether row and column headings should print, when present.</summary>
        public bool? PrintHeadings { get; internal set; }

        /// <summary>Gets whether the sheet should be centered horizontally when printed, when present.</summary>
        public bool? HorizontalCentered { get; internal set; }

        /// <summary>Gets whether the sheet should be centered vertically when printed, when present.</summary>
        public bool? VerticalCentered { get; internal set; }

        /// <summary>Gets whether fit-to-page scaling should be enabled, when present.</summary>
        public bool? FitToPage { get; internal set; }

        /// <summary>Gets whether the sheet should be printed in landscape orientation, when present.</summary>
        public bool? Landscape { get; internal set; }

        /// <summary>Gets the print scale percentage, when present.</summary>
        public ushort? Scale { get; internal set; }

        /// <summary>Gets the number of pages to fit horizontally, when present.</summary>
        public ushort? FitToWidth { get; internal set; }

        /// <summary>Gets the number of pages to fit vertically, when present.</summary>
        public ushort? FitToHeight { get; internal set; }

        /// <summary>Gets the worksheet page order for multi-page printing, when present.</summary>
        public ExcelPageOrder? PageOrder { get; internal set; }

        /// <summary>Gets the worksheet printed-size mode, when specified by a PrintSize record.</summary>
        public ushort? PrintedSize { get; internal set; }
    }
}
