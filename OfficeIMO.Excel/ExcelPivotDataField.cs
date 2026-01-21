using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes a data field in a pivot table (field name + aggregation).
    /// </summary>
    public sealed class ExcelPivotDataField {
        /// <summary>
        /// Creates a pivot data field definition.
        /// </summary>
        /// <param name="fieldName">Header name from the source range.</param>
        /// <param name="function">Aggregation function (Sum, Count, Average, ...).</param>
        /// <param name="displayName">Optional display name for the data field.</param>
        /// <param name="numberFormatId">Optional number format id to apply to the data field.</param>
        public ExcelPivotDataField(string fieldName,
            DataConsolidateFunctionValues? function = null,
            string? displayName = null,
            uint? numberFormatId = null) {
            FieldName = fieldName ?? throw new ArgumentNullException(nameof(fieldName));
            Function = function ?? DataConsolidateFunctionValues.Sum;
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
        /// Gets the optional display name for the data field.
        /// </summary>
        public string? DisplayName { get; }

        /// <summary>
        /// Gets the optional number format id for the data field.
        /// </summary>
        public uint? NumberFormatId { get; }
    }
}
