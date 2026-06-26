namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the workbook source shape carried by a legacy XLS DConRef record.
    /// </summary>
    public enum LegacyXlsDataConsolidationSourceKind {
        /// <summary>The source string did not carry a recognized DConFile prefix.</summary>
        Unknown,

        /// <summary>The source string points at an external virtual path.</summary>
        ExternalVirtualPath,

        /// <summary>The source string points at a sheet in the current workbook.</summary>
        SelfReference
    }
}
