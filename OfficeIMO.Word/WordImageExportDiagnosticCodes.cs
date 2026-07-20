namespace OfficeIMO.Word {
    /// <summary>
    /// Stable diagnostic codes emitted by Word image export.
    /// </summary>
    public static class WordImageExportDiagnosticCodes {
        /// <summary>A floating image wrap mode was rendered with reduced fidelity.</summary>
        public const string LimitedFloatingImageWrap = "limited-word-floating-image-wrap";

        /// <summary>A floating shape wrap mode was rendered with reduced fidelity.</summary>
        public const string LimitedFloatingShapeWrap = "limited-word-floating-shape-wrap";

        /// <summary>A floating text box wrap mode was rendered with reduced fidelity.</summary>
        public const string LimitedFloatingTextBoxWrap = "limited-word-floating-textbox-wrap";

        /// <summary>A SmartArt diagram was rendered with reduced fidelity.</summary>
        public const string LimitedSmartArt = "limited-word-smartart";

        /// <summary>An unsupported body element was omitted.</summary>
        public const string UnsupportedBodyElement = "unsupported-word-body-element";

        /// <summary>An external image was omitted.</summary>
        public const string UnsupportedExternalImage = "unsupported-word-external-image";

        /// <summary>A floating image was omitted.</summary>
        public const string UnsupportedFloatingImage = "unsupported-word-floating-image";

        /// <summary>An unsupported footer element was omitted.</summary>
        public const string UnsupportedFooterElement = "unsupported-word-footer-element";

        /// <summary>Footer content that overflowed its drawing area was omitted.</summary>
        public const string UnsupportedFooterOverflow = "unsupported-word-footer-overflow";

        /// <summary>Header or footer content that overflowed measurement was omitted.</summary>
        public const string UnsupportedHeaderFooterMeasurementOverflow = "unsupported-word-header-footer-measurement-overflow";

        /// <summary>An unsupported header element was omitted.</summary>
        public const string UnsupportedHeaderElement = "unsupported-word-header-element";

        /// <summary>Header content that overflowed its drawing area was omitted.</summary>
        public const string UnsupportedHeaderOverflow = "unsupported-word-header-overflow";

        /// <summary>An image was omitted.</summary>
        public const string UnsupportedImage = "unsupported-word-image";

        /// <summary>Content that overflowed keep-with-next measurement was omitted.</summary>
        public const string UnsupportedKeepMeasurementOverflow = "unsupported-word-keep-measurement-overflow";

        /// <summary>Nested-table content that overflowed its drawing area was omitted.</summary>
        public const string UnsupportedNestedTableOverflow = "unsupported-word-nested-table-overflow";

        /// <summary>A nested table was omitted.</summary>
        public const string UnsupportedNestedTable = "unsupported-word-nested-table";

        /// <summary>The requested page was outside the estimated document content.</summary>
        public const string UnsupportedPageIndex = "unsupported-word-page-index";

        /// <summary>Content that overflowed pagination was omitted.</summary>
        public const string UnsupportedPagination = "unsupported-word-pagination";

        /// <summary>A shape or one of its visual features was omitted.</summary>
        public const string UnsupportedShape = "unsupported-word-shape";

        /// <summary>A SmartArt diagram was omitted.</summary>
        public const string UnsupportedSmartArt = "unsupported-word-smartart";

        /// <summary>Table-cell content that overflowed measurement was omitted.</summary>
        public const string UnsupportedTableCellMeasurementOverflow = "unsupported-word-table-cell-measurement-overflow";

        /// <summary>Table-cell text that overflowed its drawing area was omitted.</summary>
        public const string UnsupportedTableCellTextOverflow = "unsupported-word-table-cell-text-overflow";

        /// <summary>A repeating table header that did not fit was omitted.</summary>
        public const string UnsupportedTableHeaderPagination = "unsupported-word-table-header-pagination";

        /// <summary>A table image that overflowed its drawing area was omitted.</summary>
        public const string UnsupportedTableImageOverflow = "unsupported-word-table-image-overflow";

        /// <summary>A table row that could not be split across pages was omitted.</summary>
        public const string UnsupportedTableRowPagination = "unsupported-word-table-row-pagination";

        /// <summary>A text box was omitted.</summary>
        public const string UnsupportedTextBox = "unsupported-word-textbox";
    }
}
