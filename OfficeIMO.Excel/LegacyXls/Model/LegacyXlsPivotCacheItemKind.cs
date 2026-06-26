namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the value kind stored in a PivotCache cache item record.
    /// </summary>
    public enum LegacyXlsPivotCacheItemKind {
        /// <summary>Empty PivotCache item from an SxNil record.</summary>
        Empty,

        /// <summary>Floating-point numeric PivotCache item from an SXNum record.</summary>
        Number,

        /// <summary>Boolean PivotCache item from an SxBool record.</summary>
        Boolean,

        /// <summary>Error PivotCache item from an SxErr record.</summary>
        Error,

        /// <summary>Signed integer PivotCache item from an SXInt record.</summary>
        Integer,

        /// <summary>String PivotCache item from an SXString record.</summary>
        String,

        /// <summary>Date/time PivotCache item from an SXDtr record.</summary>
        DateTime
    }
}
