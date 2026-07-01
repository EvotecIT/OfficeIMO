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

        /// <summary>Worksheet pane metadata from a Pane record.</summary>
        Pane,

        /// <summary>VBA sheet object name from a CodeName record.</summary>
        CodeName,

        /// <summary>Opaque worksheet printer settings from a Pls record.</summary>
        PrinterSettings,

        /// <summary>Worksheet printed-size mode from a PrintSize record.</summary>
        PrintSize,

        /// <summary>Worksheet sort dialog metadata from a Sort record.</summary>
        Sort,

        /// <summary>Worksheet recalculation-needed metadata from an Uncalced record.</summary>
        Uncalced,

        /// <summary>Worksheet-level phonetic display defaults from a PhoneticInfo record.</summary>
        PhoneticSettings,

        /// <summary>Worksheet hyperlink companion metadata such as HLinkTooltip records.</summary>
        Hyperlink,

        /// <summary>First/even-page header and footer metadata from a HeaderFooter record.</summary>
        HeaderFooter,

        /// <summary>Ignored formula error metadata from ISFFEC2 shared-feature records.</summary>
        IgnoredErrors,

        /// <summary>Watched-cell metadata from CellWatch records.</summary>
        CellWatches,

        /// <summary>Worksheet-level data-consolidation settings from a DCon record.</summary>
        DataConsolidation,

        /// <summary>Worksheet scenario metadata from ScenMan and SCENARIO records.</summary>
        Scenarios,

        /// <summary>Preserve-only extended metadata from a future metadata record in the sheet substream.</summary>
        FutureMetadata
    }
}
