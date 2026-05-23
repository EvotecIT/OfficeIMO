namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes a calculated pivot cache field found in an existing pivot table.
    /// </summary>
    public sealed class ExcelPivotCalculatedFieldInfo {
        /// <summary>
        /// Creates calculated pivot field readback information.
        /// </summary>
        public ExcelPivotCalculatedFieldInfo(string name, string formula, string? caption = null, uint? numberFormatId = null, string? numberFormatCode = null) {
            Name = name;
            Formula = formula;
            Caption = caption;
            NumberFormatId = numberFormatId;
            NumberFormatCode = numberFormatCode;
        }

        /// <summary>Gets the calculated field name.</summary>
        public string Name { get; }

        /// <summary>Gets the calculated field formula.</summary>
        public string Formula { get; }

        /// <summary>Gets the optional display caption.</summary>
        public string? Caption { get; }

        /// <summary>Gets the optional built-in/custom number format id.</summary>
        public uint? NumberFormatId { get; }

        /// <summary>Gets the custom number format code, when it can be resolved from workbook styles.</summary>
        public string? NumberFormatCode { get; }
    }
}
