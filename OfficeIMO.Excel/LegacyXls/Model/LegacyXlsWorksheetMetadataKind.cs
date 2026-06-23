namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies worksheet-level BIFF metadata that has been decoded into the legacy import model.
    /// </summary>
    public enum LegacyXlsWorksheetMetadataKind {
        /// <summary>Worksheet option flags from a WsBool record.</summary>
        SheetOptions,

        /// <summary>Outline gutter levels from a Guts record.</summary>
        OutlineLevels,

        /// <summary>GridSet print-gridline state metadata.</summary>
        GridSet,

        /// <summary>Row block lookup metadata from an Index record.</summary>
        RowBlockIndex,

        /// <summary>Selection metadata from a Selection record.</summary>
        Selection,

        /// <summary>VBA sheet object name from a CodeName record.</summary>
        CodeName,

        /// <summary>Opaque worksheet printer settings from a Pls record.</summary>
        PrinterSettings,

        /// <summary>Worksheet printed-size mode from a PrintSize record.</summary>
        PrintSize,

        /// <summary>Worksheet sort dialog metadata from a Sort record.</summary>
        Sort
    }
}
