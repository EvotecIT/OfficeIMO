namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes the BoundSheet visibility state of a legacy XLS sheet entry.
    /// </summary>
    public enum LegacyXlsSheetVisibility {
        /// <summary>The sheet is visible.</summary>
        Visible = 0x00,

        /// <summary>The sheet is hidden but can be unhidden by users.</summary>
        Hidden = 0x01,

        /// <summary>The sheet is hidden and can only be unhidden through VBA or compatible automation.</summary>
        VeryHidden = 0x02
    }
}
