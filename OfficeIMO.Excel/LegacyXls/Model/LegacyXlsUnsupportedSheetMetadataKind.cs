namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies metadata decoded from unsupported legacy sheet substreams.
    /// </summary>
    public enum LegacyXlsUnsupportedSheetMetadataKind {
        /// <summary>Chart printed-size mode from a PrintSize record.</summary>
        ChartPrintSize,

        /// <summary>Chart text object marker from a TxO record.</summary>
        ChartTextObject,

        /// <summary>Preserve-only extended metadata from a future metadata record in the sheet substream.</summary>
        FutureMetadata
    }
}
