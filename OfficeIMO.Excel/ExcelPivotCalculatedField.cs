using System;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes a formula-backed pivot cache field to author.
    /// </summary>
    public sealed class ExcelPivotCalculatedField {
        /// <summary>
        /// Creates a calculated pivot field.
        /// </summary>
        public ExcelPivotCalculatedField(string name, string formula, string? caption = null, uint? numberFormatId = null, string? numberFormat = null) {
            Name = string.IsNullOrWhiteSpace(name) ? throw new ArgumentNullException(nameof(name)) : name.Trim();
            Formula = string.IsNullOrWhiteSpace(formula) ? throw new ArgumentNullException(nameof(formula)) : formula.Trim();
            Caption = string.IsNullOrWhiteSpace(caption) ? null : caption!.Trim();
            NumberFormatId = numberFormatId;
            NumberFormat = string.IsNullOrWhiteSpace(numberFormat) ? null : numberFormat!.Trim();
        }

        /// <summary>Gets the calculated field name.</summary>
        public string Name { get; }

        /// <summary>Gets the calculated field formula.</summary>
        public string Formula { get; }

        /// <summary>Gets the optional display caption.</summary>
        public string? Caption { get; }

        /// <summary>Gets the optional built-in/custom number format id.</summary>
        public uint? NumberFormatId { get; }

        /// <summary>Gets the optional number format code to add to the workbook styles.</summary>
        public string? NumberFormat { get; }
    }
}
