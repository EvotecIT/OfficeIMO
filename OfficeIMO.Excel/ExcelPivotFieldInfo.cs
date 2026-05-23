using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes a source field in an existing pivot table.
    /// </summary>
    public sealed class ExcelPivotFieldInfo {
        /// <summary>
        /// Creates pivot field readback information.
        /// </summary>
        public ExcelPivotFieldInfo(
            string fieldName,
            PivotTableAxisValues? axis = null,
            FieldSortValues? sortType = null,
            uint? numberFormatId = null,
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
            IReadOnlyList<string>? hiddenItems = null,
            string? selectedItem = null,
            IReadOnlyList<string>? visibleItems = null,
            string? numberFormatCode = null) {
            FieldName = fieldName;
            Axis = axis;
            SortType = sortType;
            NumberFormatId = numberFormatId;
            NumberFormatCode = numberFormatCode;
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
            HiddenItems = hiddenItems ?? Array.Empty<string>();
            SelectedItem = selectedItem;
            VisibleItems = visibleItems ?? Array.Empty<string>();
        }

        /// <summary>
        /// Gets the source field name.
        /// </summary>
        public string FieldName { get; }

        /// <summary>
        /// Gets the pivot axis this field is assigned to, if any.
        /// </summary>
        public PivotTableAxisValues? Axis { get; }

        /// <summary>
        /// Gets the sort mode.
        /// </summary>
        public FieldSortValues? SortType { get; }

        /// <summary>
        /// Gets the field number format id.
        /// </summary>
        public uint? NumberFormatId { get; }

        /// <summary>
        /// Gets the custom number format code for the field, when it can be resolved from workbook styles.
        /// </summary>
        public string? NumberFormatCode { get; }

        /// <summary>
        /// Gets whether all items should be shown.
        /// </summary>
        public bool? ShowAll { get; }

        /// <summary>
        /// Gets whether the default subtotal is used.
        /// </summary>
        public bool? DefaultSubtotal { get; }

        /// <summary>
        /// Gets whether subtotals are shown at the top.
        /// </summary>
        public bool? SubtotalTop { get; }

        /// <summary>
        /// Gets whether a blank row is inserted after each item.
        /// </summary>
        public bool? InsertBlankRow { get; }

        /// <summary>
        /// Gets whether a page break is inserted after each item.
        /// </summary>
        public bool? InsertPageBreak { get; }

        /// <summary>
        /// Gets whether compact field layout is enabled.
        /// </summary>
        public bool? Compact { get; }

        /// <summary>
        /// Gets whether outline field layout is enabled.
        /// </summary>
        public bool? Outline { get; }

        /// <summary>
        /// Gets whether filter drop-downs are shown for the field.
        /// </summary>
        public bool? ShowDropDowns { get; }

        /// <summary>
        /// Gets whether multiple item selection is allowed.
        /// </summary>
        public bool? MultipleItemSelectionAllowed { get; }

        /// <summary>
        /// Gets whether new items are included in the filter by default.
        /// </summary>
        public bool? IncludeNewItemsInFilter { get; }

        /// <summary>
        /// Gets the custom subtotal caption.
        /// </summary>
        public string? SubtotalCaption { get; }

        /// <summary>
        /// Gets hidden item captions captured from the field item list.
        /// </summary>
        public IReadOnlyList<string> HiddenItems { get; }

        /// <summary>
        /// Gets the selected item caption when this field is used as a page/filter field.
        /// </summary>
        public string? SelectedItem { get; }

        /// <summary>
        /// Gets visible item captions captured from the field item list.
        /// </summary>
        public IReadOnlyList<string> VisibleItems { get; }
    }
}
