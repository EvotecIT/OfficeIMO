using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Identifies the SpreadsheetML surface that owns a conditional-formatting rule.
    /// </summary>
    public enum ExcelConditionalFormattingSource {
        /// <summary>The rule is stored in the standard worksheet conditional-formatting collection.</summary>
        Standard,
        /// <summary>The rule is stored in the Office 2010 conditional-formatting extension collection.</summary>
        Office2010Extension
    }

    /// <summary>
    /// Format-neutral conditional-formatting rule model. Instances returned by a worksheet
    /// are editable snapshots and can be applied through <see cref="ExcelSheet.UpdateConditionalFormattingRule"/>.
    /// </summary>
    public sealed class ExcelConditionalFormattingInfo {
        /// <summary>Gets the SpreadsheetML surface that owns the rule.</summary>
        public ExcelConditionalFormattingSource Source { get; set; }
        /// <summary>Gets the persistent extension rule identifier when the imported rule provides one.</summary>
        public string? ExtensionId { get; internal set; }
        /// <summary>Gets whether markup not projected by this model is attached to the rule or its container and will be preserved by lifecycle edits.</summary>
        public bool HasPreservedUnknownMarkup { get; internal set; }
        /// <summary>Gets or sets the A1 range covered by the rule.</summary>
        public string Range { get; set; } = string.Empty;
        /// <summary>Gets or sets the OpenXML conditional formatting type.</summary>
        public string Type { get; set; } = string.Empty;
        /// <summary>Gets or sets the OpenXML conditional formatting operator.</summary>
        public string? Operator { get; set; }
        /// <summary>Gets or sets the text payload attached to a text conditional-formatting rule.</summary>
        public string? Text { get; set; }
        /// <summary>Gets or sets the relative time period attached to a time-period conditional-formatting rule.</summary>
        public string? TimePeriod { get; set; }
        /// <summary>Gets or sets the rule priority.</summary>
        public int Priority { get; set; }
        /// <summary>Gets or sets whether evaluation stops when the rule is true.</summary>
        public bool StopIfTrue { get; set; }
        /// <summary>Gets or sets the differential format id attached to the rule, when present.</summary>
        public uint? DifferentialFormatId { get; set; }
        /// <summary>Gets or sets the resolved ARGB fill color from the rule's differential format, when present.</summary>
        public string? DifferentialFillColorArgb { get; set; }
        /// <summary>Gets or sets the resolved ARGB font color from the rule's differential format, when present.</summary>
        public string? DifferentialFontColorArgb { get; set; }
        /// <summary>Gets or sets whether the rule's differential format requests bold text, when present.</summary>
        public bool? DifferentialFontBold { get; set; }
        /// <summary>Gets or sets whether the rule's differential format requests italic text, when present.</summary>
        public bool? DifferentialFontItalic { get; set; }
        /// <summary>Gets or sets whether the rule's differential format requests underlined text, when present.</summary>
        public bool? DifferentialFontUnderline { get; set; }
        /// <summary>Gets or sets the rule's differential format font family, when present.</summary>
        public string? DifferentialFontName { get; set; }
        /// <summary>Gets or sets the rule's differential format font size in points, when present.</summary>
        public double? DifferentialFontSize { get; set; }
        /// <summary>Gets or sets the rule's differential format border, when present.</summary>
        public ExcelCellBorderSnapshot? DifferentialBorder { get; set; }
        /// <summary>Gets or sets formulas attached to the rule.</summary>
        public IReadOnlyList<string> Formulas { get; set; } = Array.Empty<string>();
        /// <summary>
        /// Gets or sets ARGB colors attached to a color-scale rule, in rule order. An empty entry in an
        /// imported rule represents a theme, indexed, or automatic color that remains preserved unless replaced.
        /// </summary>
        public IReadOnlyList<string> ColorScaleColors { get; set; } = Array.Empty<string>();
        /// <summary>Gets or sets color-scale thresholds in rule order.</summary>
        public IReadOnlyList<ExcelConditionalFormatThreshold> ColorScaleThresholds { get; set; } = Array.Empty<ExcelConditionalFormatThreshold>();
        /// <summary>Gets or sets the ARGB color attached to a data-bar rule.</summary>
        public string? DataBarColor { get; set; }
        /// <summary>Gets or sets data-bar thresholds in rule order.</summary>
        public IReadOnlyList<ExcelConditionalFormatThreshold> DataBarThresholds { get; set; } = Array.Empty<ExcelConditionalFormatThreshold>();
        /// <summary>Gets or sets whether the data-bar rule displays the underlying cell value.</summary>
        public bool DataBarShowValue { get; set; } = true;
        /// <summary>Gets or sets the minimum data-bar length percentage for Office extension rules.</summary>
        public uint? DataBarMinimumLength { get; set; }
        /// <summary>Gets or sets the maximum data-bar length percentage for Office extension rules.</summary>
        public uint? DataBarMaximumLength { get; set; }
        /// <summary>Gets or sets whether an Office extension data bar has a border.</summary>
        public bool? DataBarBorder { get; set; }
        /// <summary>Gets or sets whether an Office extension data bar uses a gradient fill.</summary>
        public bool? DataBarGradient { get; set; }
        /// <summary>Gets or sets the Office extension data-bar direction.</summary>
        public string? DataBarDirection { get; set; }
        /// <summary>Gets or sets the Office extension data-bar axis position.</summary>
        public string? DataBarAxisPosition { get; set; }
        /// <summary>Gets or sets whether negative bars reuse the positive fill color.</summary>
        public bool? DataBarNegativeColorSameAsPositive { get; set; }
        /// <summary>Gets or sets whether negative bar borders reuse the positive border color.</summary>
        public bool? DataBarNegativeBorderColorSameAsPositive { get; set; }
        /// <summary>Gets or sets the ARGB data-bar border color.</summary>
        public string? DataBarBorderColor { get; set; }
        /// <summary>Gets or sets the ARGB negative data-bar fill color.</summary>
        public string? DataBarNegativeColor { get; set; }
        /// <summary>Gets or sets the ARGB negative data-bar border color.</summary>
        public string? DataBarNegativeBorderColor { get; set; }
        /// <summary>Gets or sets the ARGB data-bar axis color.</summary>
        public string? DataBarAxisColor { get; set; }
        /// <summary>Gets or sets the icon-set name attached to an icon-set rule.</summary>
        public string? IconSet { get; set; }
        /// <summary>Gets or sets whether the icon-set rule displays the underlying cell value.</summary>
        public bool IconSetShowValue { get; set; } = true;
        /// <summary>Gets or sets whether the icon-set rule reverses icon order.</summary>
        public bool IconSetReverse { get; set; }
        /// <summary>Gets or sets whether an Office extension icon set interprets thresholds as percentages.</summary>
        public bool? IconSetPercent { get; set; }
        /// <summary>Gets or sets whether an Office extension icon set uses explicit custom icons.</summary>
        public bool? IconSetCustom { get; set; }
        /// <summary>Gets or sets icon-set thresholds in rule order.</summary>
        public IReadOnlyList<ExcelConditionalIconSetThreshold> IconSetThresholds { get; set; } = Array.Empty<ExcelConditionalIconSetThreshold>();
        /// <summary>Gets or sets custom icon references in rule order.</summary>
        public IReadOnlyList<ExcelConditionalFormatIcon> CustomIcons { get; set; } = Array.Empty<ExcelConditionalFormatIcon>();
        /// <summary>Gets or sets the top/bottom rule rank, when present.</summary>
        public uint? TopBottomRank { get; set; }
        /// <summary>Gets or sets whether the top/bottom rule selects bottom values.</summary>
        public bool TopBottomBottom { get; set; }
        /// <summary>Gets or sets whether the top/bottom rule rank is a percentage.</summary>
        public bool TopBottomPercent { get; set; }
        /// <summary>Gets or sets whether the above-average rule selects values above the average.</summary>
        public bool AboveAverageAbove { get; set; } = true;
        /// <summary>Gets or sets whether the above-average rule includes values equal to the average.</summary>
        public bool AboveAverageEqual { get; set; }
        /// <summary>Gets or sets the standard-deviation threshold for above-average rules, when present.</summary>
        public int? AboveAverageStdDev { get; set; }

        internal ExcelSheet? OwnerSheet { get; set; }
        internal OpenXmlElement? OwnerContainer { get; set; }
        internal OpenXmlElement? BackingRule { get; set; }
        internal string? ProjectedFormulaSignature { get; set; }
        internal string? ProjectedVisualSignature { get; set; }
        internal string? ProjectedStyleSignature { get; set; }
        internal string? ProjectedType { get; set; }
    }

    /// <summary>
    /// Threshold metadata for a conditional-formatting value object.
    /// </summary>
    public sealed class ExcelConditionalFormatThreshold {
        /// <summary>Gets or sets the threshold type, such as Min, Max, Number, Percent, Percentile, or Formula.</summary>
        public string Type { get; set; } = string.Empty;
        /// <summary>Gets or sets the raw threshold value, when present.</summary>
        public string? Value { get; set; }
    }

    /// <summary>
    /// Threshold metadata for a conditional-formatting icon-set rule.
    /// </summary>
    public sealed class ExcelConditionalIconSetThreshold {
        /// <summary>Gets or sets the threshold type, such as Percent, Number, Minimum, or Maximum.</summary>
        public string Type { get; set; } = string.Empty;
        /// <summary>Gets or sets the raw threshold value, when present.</summary>
        public string? Value { get; set; }
        /// <summary>Gets or sets whether values equal to the threshold are included in the higher icon bucket.</summary>
        public bool GreaterThanOrEqual { get; set; } = true;
    }

    /// <summary>
    /// Identifies an icon used by a custom Office extension icon-set rule.
    /// </summary>
    public sealed class ExcelConditionalFormatIcon {
        /// <summary>Gets or sets the icon-set family.</summary>
        public string IconSet { get; set; } = string.Empty;
        /// <summary>Gets or sets the zero-based icon identifier within the family.</summary>
        public uint IconId { get; set; }
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
        internal bool PreserveShowMessageFlags { get; set; }
        /// <summary>Gets or sets the validation error alert style.</summary>
        public ExcelDataValidationErrorStyle? ErrorStyle { get; set; }
        /// <summary>Gets or sets whether Excel should hide the in-cell dropdown for list validations. Leave null to preserve the existing value.</summary>
        public bool? SuppressDropDown { get; set; }
    }
}
