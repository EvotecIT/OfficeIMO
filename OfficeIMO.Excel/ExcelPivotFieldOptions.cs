using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes formatting and display options for a pivot source field.
    /// </summary>
    public sealed class ExcelPivotFieldOptions {
        /// <summary>
        /// Creates pivot field options for a source field.
        /// </summary>
        /// <param name="fieldName">Header name from the source range.</param>
        /// <param name="sortType">Optional pivot field sort mode.</param>
        /// <param name="numberFormatId">Optional number format id to apply to the pivot field.</param>
        /// <param name="numberFormat">Optional number format code to apply to the pivot field.</param>
        /// <param name="showAll">Whether to show all items for the pivot field.</param>
        /// <param name="defaultSubtotal">Whether to use the default subtotal for the pivot field.</param>
        /// <param name="subtotalTop">Whether subtotals should be shown at the top.</param>
        /// <param name="insertBlankRow">Whether to insert a blank row after each item.</param>
        public ExcelPivotFieldOptions(
            string fieldName,
            FieldSortValues? sortType = null,
            uint? numberFormatId = null,
            string? numberFormat = null,
            bool? showAll = null,
            bool? defaultSubtotal = null,
            bool? subtotalTop = null,
            bool? insertBlankRow = null) {
            FieldName = fieldName ?? throw new ArgumentNullException(nameof(fieldName));
            SortType = sortType;
            NumberFormatId = numberFormatId;
            NumberFormat = numberFormat;
            ShowAll = showAll;
            DefaultSubtotal = defaultSubtotal;
            SubtotalTop = subtotalTop;
            InsertBlankRow = insertBlankRow;
        }

        /// <summary>
        /// Gets the source field name.
        /// </summary>
        public string FieldName { get; }

        /// <summary>
        /// Gets the optional pivot field sort mode.
        /// </summary>
        public FieldSortValues? SortType { get; }

        /// <summary>
        /// Gets the optional number format id.
        /// </summary>
        public uint? NumberFormatId { get; }

        /// <summary>
        /// Gets the optional number format code.
        /// </summary>
        public string? NumberFormat { get; }

        /// <summary>
        /// Gets whether all items should be shown for the field.
        /// </summary>
        public bool? ShowAll { get; }

        /// <summary>
        /// Gets whether the default subtotal should be used.
        /// </summary>
        public bool? DefaultSubtotal { get; }

        /// <summary>
        /// Gets whether subtotals should be shown at the top.
        /// </summary>
        public bool? SubtotalTop { get; }

        /// <summary>
        /// Gets whether to insert a blank row after each item.
        /// </summary>
        public bool? InsertBlankRow { get; }
    }
}
