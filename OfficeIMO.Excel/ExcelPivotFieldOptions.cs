using System;
using System.Collections.Generic;
using System.Linq;
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
        /// <param name="insertPageBreak">Whether to insert a page break after each item.</param>
        /// <param name="compact">Whether compact layout is enabled for this field.</param>
        /// <param name="outline">Whether outline layout is enabled for this field.</param>
        /// <param name="showDropDowns">Whether field drop-downs are shown.</param>
        /// <param name="multipleItemSelectionAllowed">Whether multiple filter item selection is allowed.</param>
        /// <param name="includeNewItemsInFilter">Whether new source items are included in the filter by default.</param>
        /// <param name="subtotalCaption">Optional custom subtotal caption.</param>
        /// <param name="hiddenItems">Optional source item captions to hide.</param>
        /// <param name="visibleItems">Optional source item captions to show; other known items are hidden.</param>
        /// <param name="selectedItem">Optional page-field item caption to select.</param>
        public ExcelPivotFieldOptions(
            string fieldName,
            FieldSortValues? sortType = null,
            uint? numberFormatId = null,
            string? numberFormat = null,
            bool? showAll = null,
            bool? defaultSubtotal = null,
            bool? subtotalTop = null,
            bool? insertBlankRow = null,
            bool? insertPageBreak = null,
            bool? compact = null,
            bool? outline = null,
            bool? showDropDowns = null,
            bool? multipleItemSelectionAllowed = null,
            bool? includeNewItemsInFilter = null,
            string? subtotalCaption = null,
            IEnumerable<string>? hiddenItems = null,
            IEnumerable<string>? visibleItems = null,
            string? selectedItem = null) {
            FieldName = fieldName ?? throw new ArgumentNullException(nameof(fieldName));
            SortType = sortType;
            NumberFormatId = numberFormatId;
            NumberFormat = numberFormat;
            ShowAll = showAll;
            DefaultSubtotal = defaultSubtotal;
            SubtotalTop = subtotalTop;
            InsertBlankRow = insertBlankRow;
            InsertPageBreak = insertPageBreak;
            Compact = compact;
            Outline = outline;
            ShowDropDowns = showDropDowns;
            MultipleItemSelectionAllowed = multipleItemSelectionAllowed;
            IncludeNewItemsInFilter = includeNewItemsInFilter;
            SubtotalCaption = subtotalCaption;
            HiddenItems = NormalizeItems(hiddenItems);
            VisibleItems = NormalizeItems(visibleItems);
            SelectedItem = string.IsNullOrWhiteSpace(selectedItem) ? null : selectedItem!.Trim();
            if (HiddenItems.Count > 0 && VisibleItems.Count > 0) {
                throw new ArgumentException("Specify either hiddenItems or visibleItems, not both.");
            }
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

        /// <summary>
        /// Gets whether to insert a page break after each item.
        /// </summary>
        public bool? InsertPageBreak { get; }

        /// <summary>
        /// Gets whether compact layout is enabled for this field.
        /// </summary>
        public bool? Compact { get; }

        /// <summary>
        /// Gets whether outline layout is enabled for this field.
        /// </summary>
        public bool? Outline { get; }

        /// <summary>
        /// Gets whether field drop-downs are shown.
        /// </summary>
        public bool? ShowDropDowns { get; }

        /// <summary>
        /// Gets whether multiple item selection is allowed.
        /// </summary>
        public bool? MultipleItemSelectionAllowed { get; }

        /// <summary>
        /// Gets whether new source items are included in the filter by default.
        /// </summary>
        public bool? IncludeNewItemsInFilter { get; }

        /// <summary>
        /// Gets the optional custom subtotal caption.
        /// </summary>
        public string? SubtotalCaption { get; }

        /// <summary>
        /// Gets source item captions to hide.
        /// </summary>
        public IReadOnlyList<string> HiddenItems { get; }

        /// <summary>
        /// Gets source item captions to show; other known items are hidden.
        /// </summary>
        public IReadOnlyList<string> VisibleItems { get; }

        /// <summary>
        /// Gets the selected page-field item caption.
        /// </summary>
        public string? SelectedItem { get; }

        private static IReadOnlyList<string> NormalizeItems(IEnumerable<string>? items) {
            if (items == null) return Array.Empty<string>();
            return items
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Select(item => item.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }
    }
}
