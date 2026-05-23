using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes a pivot filter found in an existing pivot table.
    /// </summary>
    public sealed class ExcelPivotFilterInfo {
        /// <summary>
        /// Creates pivot filter readback information.
        /// </summary>
        public ExcelPivotFilterInfo(
            string fieldName,
            PivotFilterValues? type,
            string? value1,
            string? value2,
            string? dataFieldName,
            string? name,
            string? description,
            bool? isTop = null,
            bool? isPercent = null,
            string? filterValue = null) {
            FieldName = fieldName;
            Type = type;
            Value1 = value1;
            Value2 = value2;
            DataFieldName = dataFieldName;
            Name = name;
            Description = description;
            IsTop = isTop;
            IsPercent = isPercent;
            FilterValue = filterValue;
        }

        /// <summary>Gets the source field name being filtered.</summary>
        public string FieldName { get; }

        /// <summary>Gets the Open XML pivot filter type.</summary>
        public PivotFilterValues? Type { get; }

        /// <summary>Gets the first filter value.</summary>
        public string? Value1 { get; }

        /// <summary>Gets the second filter value for between-style filters.</summary>
        public string? Value2 { get; }

        /// <summary>Gets the data field used by value filters.</summary>
        public string? DataFieldName { get; }

        /// <summary>Gets the filter display name.</summary>
        public string? Name { get; }

        /// <summary>Gets the filter description.</summary>
        public string? Description { get; }

        /// <summary>Gets whether a top/bottom filter keeps top values. False means bottom values.</summary>
        public bool? IsTop { get; }

        /// <summary>Gets whether a top/bottom filter uses a percentage threshold.</summary>
        public bool? IsPercent { get; }

        /// <summary>Gets the optional calculated top/bottom filter value threshold.</summary>
        public string? FilterValue { get; }
    }
}
