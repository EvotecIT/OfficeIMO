namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies a preserve-only compound container feature discovered outside the Workbook BIFF stream.
    /// </summary>
    public enum LegacyXlsCompoundFeatureRecordKind {
        /// <summary>VBA project storage was found in the XLS compound container.</summary>
        VbaProject,

        /// <summary>Embedded OLE object storage was found in the XLS compound container.</summary>
        OleObject
    }
}
