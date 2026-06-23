namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies a shallow preserve-only chart BIFF record category.
    /// </summary>
    public enum LegacyXlsChartRecordKind {
        /// <summary>Chart record that is currently preserved without record-specific field decoding.</summary>
        PreserveOnly,

        /// <summary>Chart container or chart stream boundary record.</summary>
        Container,

        /// <summary>Chart data series record.</summary>
        Series,

        /// <summary>Chart axis, tick, or value range record.</summary>
        Axis,

        /// <summary>Chart text, label, or font record.</summary>
        Text,

        /// <summary>Chart line, area, marker, or data format record.</summary>
        Formatting,

        /// <summary>Chart positioning, frame, or plot-area layout record.</summary>
        Layout,

        /// <summary>Chart type or type-specific option record.</summary>
        ChartType,

        /// <summary>Chart extension record.</summary>
        Extension
    }
}
