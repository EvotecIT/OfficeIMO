namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes the printed-size mode of a legacy XLS chart sheet PrintSize record.
    /// </summary>
    public enum LegacyXlsChartPrintSize {
        /// <summary>Print settings are unchanged from workbook defaults in a UserSViewBegin block.</summary>
        DefaultsUnchanged = 0x0000,

        /// <summary>The chart is resized to fill the page without preserving original chart proportions.</summary>
        FillPage = 0x0001,

        /// <summary>The chart is resized proportionally to fit the page.</summary>
        FitPage = 0x0002,

        /// <summary>The printed chart size is defined by the Chart record.</summary>
        DefinedInChartRecord = 0x0003
    }
}
