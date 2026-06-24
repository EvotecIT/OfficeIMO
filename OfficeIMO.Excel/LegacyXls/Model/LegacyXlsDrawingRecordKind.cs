namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies a shallow preserve-only drawing or object BIFF record category.
    /// </summary>
    public enum LegacyXlsDrawingRecordKind {
        /// <summary>Drawing record that is currently preserved without record-specific field decoding.</summary>
        PreserveOnly,

        /// <summary>Workbook drawing group container record.</summary>
        DrawingGroup,

        /// <summary>Worksheet drawing container record.</summary>
        Drawing,

        /// <summary>Worksheet object record.</summary>
        Object,

        /// <summary>Text object record.</summary>
        TextObject,

        /// <summary>Drawing shape-properties future-record stream.</summary>
        ShapePropertiesStream,

        /// <summary>Drawing text-properties future-record stream.</summary>
        TextPropertiesStream,

        /// <summary>Drawing rich-text future-record stream.</summary>
        RichTextStream
    }
}
