using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Read-only snapshot of a conditional formatting rule.
    /// </summary>
    public sealed class ExcelConditionalFormattingInfo {
        /// <summary>Gets or sets the A1 range covered by the rule.</summary>
        public string Range { get; set; } = string.Empty;
        /// <summary>Gets or sets the OpenXML conditional formatting type.</summary>
        public string Type { get; set; } = string.Empty;
        /// <summary>Gets or sets the OpenXML conditional formatting operator.</summary>
        public string? Operator { get; set; }
        /// <summary>Gets or sets the rule priority.</summary>
        public int Priority { get; set; }
        /// <summary>Gets or sets whether evaluation stops when the rule is true.</summary>
        public bool StopIfTrue { get; set; }
        /// <summary>Gets or sets formulas attached to the rule.</summary>
        public IReadOnlyList<string> Formulas { get; set; } = Array.Empty<string>();
        /// <summary>Gets or sets ARGB colors attached to a color-scale rule, in rule order.</summary>
        public IReadOnlyList<string> ColorScaleColors { get; set; } = Array.Empty<string>();
        /// <summary>Gets or sets the ARGB color attached to a data-bar rule.</summary>
        public string? DataBarColor { get; set; }
        /// <summary>Gets or sets the icon-set name attached to an icon-set rule.</summary>
        public string? IconSet { get; set; }
        /// <summary>Gets or sets whether the icon-set rule displays the underlying cell value.</summary>
        public bool IconSetShowValue { get; set; } = true;
        /// <summary>Gets or sets whether the icon-set rule reverses icon order.</summary>
        public bool IconSetReverse { get; set; }
    }

    /// <summary>
    /// Read-only snapshot of a data validation rule.
    /// </summary>
    public sealed class ExcelDataValidationInfo {
        /// <summary>Gets or sets the A1 range covered by the validation.</summary>
        public string Range { get; set; } = string.Empty;
        /// <summary>Gets or sets the OpenXML validation type.</summary>
        public string Type { get; set; } = string.Empty;
        /// <summary>Gets or sets the OpenXML validation operator.</summary>
        public string? Operator { get; set; }
        /// <summary>Gets or sets whether blank values are allowed.</summary>
        public bool AllowBlank { get; set; }
        /// <summary>Gets or sets whether Excel should hide the in-cell dropdown for list validations.</summary>
        public bool SuppressDropDown { get; set; }
        /// <summary>Gets or sets the OpenXML validation error style.</summary>
        public string? ErrorStyle { get; set; }
        /// <summary>Gets or sets whether Excel should show the input prompt.</summary>
        public bool ShowInputMessage { get; set; }
        /// <summary>Gets or sets whether Excel should show the validation error.</summary>
        public bool ShowErrorMessage { get; set; }
        /// <summary>Gets or sets the first validation formula.</summary>
        public string? Formula1 { get; set; }
        /// <summary>Gets or sets the second validation formula.</summary>
        public string? Formula2 { get; set; }
        /// <summary>Gets or sets the input prompt title.</summary>
        public string? PromptTitle { get; set; }
        /// <summary>Gets or sets the input prompt text.</summary>
        public string? Prompt { get; set; }
        /// <summary>Gets or sets the error title.</summary>
        public string? ErrorTitle { get; set; }
        /// <summary>Gets or sets the error text.</summary>
        public string? Error { get; set; }
    }

    /// <summary>
    /// User-facing prompt/error options for data validation.
    /// </summary>
    public sealed class ExcelDataValidationMessageOptions {
        /// <summary>Gets or sets the input prompt title.</summary>
        public string? PromptTitle { get; set; }
        /// <summary>Gets or sets the input prompt text.</summary>
        public string? Prompt { get; set; }
        /// <summary>Gets or sets the validation error title.</summary>
        public string? ErrorTitle { get; set; }
        /// <summary>Gets or sets the validation error text.</summary>
        public string? Error { get; set; }
        /// <summary>Gets or sets whether Excel should show the input prompt.</summary>
        public bool ShowInputMessage { get; set; }
        /// <summary>Gets or sets whether Excel should show the validation error.</summary>
        public bool ShowErrorMessage { get; set; }
        /// <summary>Gets or sets the validation error alert style.</summary>
        public DataValidationErrorStyleValues? ErrorStyle { get; set; }
        /// <summary>Gets or sets whether Excel should hide the in-cell dropdown for list validations. Leave null to preserve the existing value.</summary>
        public bool? SuppressDropDown { get; set; }
    }
}
