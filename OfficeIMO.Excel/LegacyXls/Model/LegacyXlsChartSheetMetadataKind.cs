namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies metadata decoded from a legacy chart-sheet substream.
    /// </summary>
    public enum LegacyXlsChartSheetMetadataKind {
        /// <summary>Chart printed-size mode from a PrintSize record.</summary>
        ChartPrintSize,

        /// <summary>Chart text object marker from a TxO record.</summary>
        ChartTextObject,

        /// <summary>Extended metadata from a future metadata record in the chart-sheet substream.</summary>
        FutureMetadata
    }
}
