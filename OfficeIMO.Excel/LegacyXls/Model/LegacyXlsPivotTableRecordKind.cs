namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies a preserve-only PivotTable BIFF record decoded into the legacy XLS import model.
    /// </summary>
    public enum LegacyXlsPivotTableRecordKind {
        /// <summary>PivotTable record that is currently preserved without record-specific field decoding.</summary>
        PreserveOnly,

        /// <summary>Data item metadata from an SXDI record.</summary>
        DataItem,

        /// <summary>Grouping range metadata from an SXRng record.</summary>
        GroupingRange,

        /// <summary>Extended pivot field metadata from an SXVDEx record.</summary>
        ExtendedPivotField
    }
}
