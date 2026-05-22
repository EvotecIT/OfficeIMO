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
        /// Creates a data field info instance.
        /// </summary>
        public ExcelPivotDataFieldInfo(string fieldName, DataConsolidateFunctionValues function, string? displayName,
            uint? numberFormatId, ShowDataAsValues? showDataAs, int? baseField, uint? baseItem) {
            FieldName = fieldName;
            Function = function;
            DisplayName = displayName;
            NumberFormatId = numberFormatId;
            ShowDataAs = showDataAs;
            BaseField = baseField;
            BaseItem = baseItem;
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

        /// <summary>
        /// Gets the show-values-as calculation mode applied to the data field.
        /// </summary>
        public ShowDataAsValues? ShowDataAs { get; }

        /// <summary>
        /// Gets the base field index for show-values-as calculations that require one.
        /// </summary>
        public int? BaseField { get; }

        /// <summary>
        /// Gets the base item index for show-values-as calculations that require one.
        /// </summary>
        public uint? BaseItem { get; }
    }
}
