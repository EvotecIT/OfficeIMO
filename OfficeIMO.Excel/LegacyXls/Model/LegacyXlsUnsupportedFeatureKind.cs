namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies unsupported or preserve-only legacy XLS features discovered during import.
    /// </summary>
    public enum LegacyXlsUnsupportedFeatureKind {
        /// <summary>
        /// A BIFF record was discovered but is not imported by the current phase.
        /// </summary>
        UnsupportedRecord,

        /// <summary>
        /// The workbook is password-to-open encrypted.
        /// </summary>
        EncryptedWorkbook,

        /// <summary>
        /// The workbook uses an older or otherwise unsupported BIFF version.
        /// </summary>
        UnsupportedBiffVersion,

        /// <summary>
        /// A macro sheet entry was discovered.
        /// </summary>
        MacroSheet,

        /// <summary>
        /// A chart sheet entry was discovered.
        /// </summary>
        ChartSheet,

        /// <summary>
        /// A VBA module sheet entry was discovered.
        /// </summary>
        VbaModuleSheet,

        /// <summary>
        /// VBA project storage was discovered in the OLE compound container.
        /// </summary>
        VbaProject,

        /// <summary>
        /// Embedded OLE object storage was discovered in the OLE compound container.
        /// </summary>
        OleObject,

        /// <summary>
        /// A dialog sheet entry was discovered.
        /// </summary>
        DialogSheet,

        /// <summary>
        /// A hyperlink record was present but its target shape is not supported yet.
        /// </summary>
        Hyperlink,

        /// <summary>
        /// A legacy comment record shape was present but not imported.
        /// </summary>
        Comment,

        /// <summary>
        /// Drawing or object records were present but not imported.
        /// </summary>
        DrawingObject,

        /// <summary>
        /// PivotTable records were present but not imported.
        /// </summary>
        PivotTable,

        /// <summary>
        /// External reference records or supporting links were present but not imported.
        /// </summary>
        ExternalReference,

        /// <summary>
        /// AutoFilter criteria or supporting records were present but not imported.
        /// </summary>
        AutoFilterCriteria,

        /// <summary>
        /// Data validation records were present but not imported.
        /// </summary>
        DataValidation,

        /// <summary>
        /// Conditional formatting records were present but not imported.
        /// </summary>
        ConditionalFormatting,

        /// <summary>
        /// Chart records were present but not imported.
        /// </summary>
        Chart,

        /// <summary>
        /// Extended style formatting records were present but not projected.
        /// </summary>
        StyleExtension,

        /// <summary>
        /// Extended table style records were present but not projected.
        /// </summary>
        TableStyle,

        /// <summary>
        /// Extended theme records were present but not projected.
        /// </summary>
        Theme,

        /// <summary>
        /// Extended workbook metadata records were present but not projected.
        /// </summary>
        WorkbookMetadata,

        /// <summary>
        /// Future feature extension records were present but not projected.
        /// </summary>
        FeatureExtension,

        /// <summary>
        /// Phonetic guide records were present but not projected.
        /// </summary>
        PhoneticGuide,

        /// <summary>
        /// An unsupported worksheet-like sheet type was discovered.
        /// </summary>
        UnsupportedSheet
    }
}
