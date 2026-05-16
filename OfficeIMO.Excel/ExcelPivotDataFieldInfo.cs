using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes a data field in an existing pivot table.
    /// </summary>
    public sealed class ExcelPivotDataFieldInfo {
        /// <summary>
        /// Creates a data field info instance.
        /// </summary>
        public ExcelPivotDataFieldInfo(string fieldName, DataConsolidateFunctionValues function, string? displayName)
            : this(fieldName, function, displayName, null) {
        }

        /// <summary>
        /// Creates a data field info instance.
        /// </summary>
        public ExcelPivotDataFieldInfo(string fieldName, DataConsolidateFunctionValues function, string? displayName, uint? numberFormatId) {
            FieldName = fieldName;
            Function = function;
            DisplayName = displayName;
            NumberFormatId = numberFormatId;
        }

        /// <summary>
        /// Gets the source field name.
        /// </summary>
        public string FieldName { get; }

        /// <summary>
        /// Gets the aggregation function.
        /// </summary>
        public DataConsolidateFunctionValues Function { get; }

        /// <summary>
        /// Gets the display name for the data field.
        /// </summary>
        public string? DisplayName { get; }

        /// <summary>
        /// Gets the number format id applied to the data field.
        /// </summary>
        public uint? NumberFormatId { get; }
    }
}
