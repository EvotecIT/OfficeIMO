using DocumentFormat.OpenXml.Spreadsheet;
using OfficeColor = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Fluent conditional formatting builder for an <see cref="ExcelRange"/>.
    /// </summary>
    public sealed class ExcelRangeConditionalFormattingBuilder {
        private readonly ExcelRange _range;

        internal ExcelRangeConditionalFormattingBuilder(ExcelRange range) {
            _range = range ?? throw new ArgumentNullException(nameof(range));
        }

        /// <summary>
        /// Applies a cell-value comparison rule.
        /// </summary>
        public ExcelRange CellIs(ConditionalFormattingOperatorValues @operator, string formula1) {
            return CellIs(@operator, formula1, null, null);
        }

        /// <summary>
        /// Applies a cell-value comparison rule.
        /// </summary>
        public ExcelRange CellIs(ConditionalFormattingOperatorValues @operator, string formula1, string? formula2) {
            return CellIs(@operator, formula1, formula2, null);
        }

        /// <summary>
        /// Applies a cell-value comparison rule.
        /// </summary>
        public ExcelRange CellIs(ConditionalFormattingOperatorValues @operator, string formula1, string? formula2 = null, string? fillColor = null) {
            _range.Sheet.AddConditionalRule(_range.Address, @operator, formula1, formula2, fillColor);
            return _range;
        }

        /// <summary>
        /// Applies a greater-than cell-value rule.
        /// </summary>
        public ExcelRange GreaterThan(string formula) {
            return GreaterThan(formula, null);
        }

        /// <summary>
        /// Applies a greater-than cell-value rule.
        /// </summary>
        public ExcelRange GreaterThan(string formula, string? fillColor) {
            return CellIs(ConditionalFormattingOperatorValues.GreaterThan, formula, null, fillColor);
        }

        /// <summary>
        /// Applies a less-than cell-value rule.
        /// </summary>
        public ExcelRange LessThan(string formula) {
            return LessThan(formula, null);
        }

        /// <summary>
        /// Applies a less-than cell-value rule.
        /// </summary>
        public ExcelRange LessThan(string formula, string? fillColor) {
            return CellIs(ConditionalFormattingOperatorValues.LessThan, formula, null, fillColor);
        }

        /// <summary>
        /// Applies a between cell-value rule.
        /// </summary>
        public ExcelRange Between(string formula1, string formula2) {
            return Between(formula1, formula2, null);
        }

        /// <summary>
        /// Applies a between cell-value rule.
        /// </summary>
        public ExcelRange Between(string formula1, string formula2, string? fillColor) {
            return CellIs(ConditionalFormattingOperatorValues.Between, formula1, formula2, fillColor);
        }

        /// <summary>
        /// Applies a formula-based rule.
        /// </summary>
        public ExcelRange Formula(string formula) {
            return Formula(formula, stopIfTrue: false, fillColor: null);
        }

        /// <summary>
        /// Applies a formula-based rule.
        /// </summary>
        public ExcelRange Formula(string formula, bool stopIfTrue) {
            return Formula(formula, stopIfTrue, fillColor: null);
        }

        /// <summary>
        /// Applies a formula-based rule.
        /// </summary>
        public ExcelRange Formula(string formula, bool stopIfTrue = false, string? fillColor = null) {
            _range.Sheet.AddConditionalFormulaRule(_range.Address, formula, stopIfTrue, fillColor);
            return _range;
        }

        /// <summary>
        /// Applies a contains-text rule.
        /// </summary>
        public ExcelRange ContainsText(string text) {
            return ContainsText(text, fillColor: null);
        }

        /// <summary>
        /// Applies a contains-text rule with an optional solid fill.
        /// </summary>
        public ExcelRange ContainsText(string text, string? fillColor) {
            _range.Sheet.AddConditionalContainsTextRule(_range.Address, text, fillColor);
            return _range;
        }

        /// <summary>
        /// Applies a not-contains-text rule.
        /// </summary>
        public ExcelRange NotContainsText(string text) {
            return NotContainsText(text, fillColor: null);
        }

        /// <summary>
        /// Applies a not-contains-text rule with an optional solid fill.
        /// </summary>
        public ExcelRange NotContainsText(string text, string? fillColor) {
            _range.Sheet.AddConditionalNotContainsTextRule(_range.Address, text, fillColor);
            return _range;
        }

        /// <summary>
        /// Applies a begins-with-text rule.
        /// </summary>
        public ExcelRange BeginsWithText(string text) {
            return BeginsWithText(text, fillColor: null);
        }

        /// <summary>
        /// Applies a begins-with-text rule with an optional solid fill.
        /// </summary>
        public ExcelRange BeginsWithText(string text, string? fillColor) {
            _range.Sheet.AddConditionalBeginsWithTextRule(_range.Address, text, fillColor);
            return _range;
        }

        /// <summary>
        /// Applies an ends-with-text rule.
        /// </summary>
        public ExcelRange EndsWithText(string text) {
            return EndsWithText(text, fillColor: null);
        }

        /// <summary>
        /// Applies an ends-with-text rule with an optional solid fill.
        /// </summary>
        public ExcelRange EndsWithText(string text, string? fillColor) {
            _range.Sheet.AddConditionalEndsWithTextRule(_range.Address, text, fillColor);
            return _range;
        }

        /// <summary>
        /// Applies a duplicate-values rule.
        /// </summary>
        public ExcelRange DuplicateValues() {
            return DuplicateValues(fillColor: null);
        }

        /// <summary>
        /// Applies a duplicate-values rule with an optional solid fill.
        /// </summary>
        public ExcelRange DuplicateValues(string? fillColor) {
            _range.Sheet.AddConditionalDuplicateValuesRule(_range.Address, fillColor);
            return _range;
        }

        /// <summary>
        /// Applies a unique-values rule.
        /// </summary>
        public ExcelRange UniqueValues() {
            return UniqueValues(fillColor: null);
        }

        /// <summary>
        /// Applies a unique-values rule with an optional solid fill.
        /// </summary>
        public ExcelRange UniqueValues(string? fillColor) {
            _range.Sheet.AddConditionalUniqueValuesRule(_range.Address, fillColor);
            return _range;
        }

        /// <summary>
        /// Applies an above-average rule.
        /// </summary>
        public ExcelRange AboveAverage() {
            return AboveAverage(fillColor: null);
        }

        /// <summary>
        /// Applies an above-average rule with an optional solid fill.
        /// </summary>
        public ExcelRange AboveAverage(string? fillColor, bool equalAverage = false) {
            _range.Sheet.AddConditionalAboveAverageRule(_range.Address, aboveAverage: true, equalAverage: equalAverage, fillColor: fillColor);
            return _range;
        }

        /// <summary>
        /// Applies a below-average rule.
        /// </summary>
        public ExcelRange BelowAverage() {
            return BelowAverage(fillColor: null);
        }

        /// <summary>
        /// Applies a below-average rule with an optional solid fill.
        /// </summary>
        public ExcelRange BelowAverage(string? fillColor, bool equalAverage = false) {
            _range.Sheet.AddConditionalAboveAverageRule(_range.Address, aboveAverage: false, equalAverage: equalAverage, fillColor: fillColor);
            return _range;
        }

        /// <summary>
        /// Applies a top-values rule.
        /// </summary>
        public ExcelRange Top(uint rank, bool percent = false) {
            return Top(rank, fillColor: null, percent: percent);
        }

        /// <summary>
        /// Applies a top-values rule with an optional solid fill.
        /// </summary>
        public ExcelRange Top(uint rank, string? fillColor, bool percent = false) {
            _range.Sheet.AddConditionalTopBottomRule(_range.Address, rank, bottom: false, percent: percent, fillColor: fillColor);
            return _range;
        }

        /// <summary>
        /// Applies a bottom-values rule.
        /// </summary>
        public ExcelRange Bottom(uint rank, bool percent = false) {
            return Bottom(rank, fillColor: null, percent: percent);
        }

        /// <summary>
        /// Applies a bottom-values rule with an optional solid fill.
        /// </summary>
        public ExcelRange Bottom(uint rank, string? fillColor, bool percent = false) {
            _range.Sheet.AddConditionalTopBottomRule(_range.Address, rank, bottom: true, percent: percent, fillColor: fillColor);
            return _range;
        }

        /// <summary>
        /// Applies an above-average rule.
        /// </summary>
        public ExcelRange AboveAverage(bool equalAverage = false, uint? standardDeviation = null) {
            _range.Sheet.AddConditionalAboveAverageRule(_range.Address, aboveAverage: true, equalAverage: equalAverage, standardDeviation: standardDeviation);
            return _range;
        }

        /// <summary>
        /// Applies a below-average rule.
        /// </summary>
        public ExcelRange BelowAverage(bool equalAverage = false, uint? standardDeviation = null) {
            _range.Sheet.AddConditionalAboveAverageRule(_range.Address, aboveAverage: false, equalAverage: equalAverage, standardDeviation: standardDeviation);
            return _range;
        }

        /// <summary>
        /// Applies a begins-with text rule.
        /// </summary>
        public ExcelRange BeginsWith(string text) {
            _range.Sheet.AddConditionalTextRule(_range.Address, ConditionalFormatValues.BeginsWith, text);
            return _range;
        }

        /// <summary>
        /// Applies an ends-with text rule.
        /// </summary>
        public ExcelRange EndsWith(string text) {
            _range.Sheet.AddConditionalTextRule(_range.Address, ConditionalFormatValues.EndsWith, text);
            return _range;
        }

        /// <summary>
        /// Applies a blanks rule.
        /// </summary>
        public ExcelRange Blanks() {
            _range.Sheet.AddConditionalBlanksRule(_range.Address, containsBlanks: true);
            return _range;
        }

        /// <summary>
        /// Applies a non-blanks rule.
        /// </summary>
        public ExcelRange NonBlanks() {
            _range.Sheet.AddConditionalBlanksRule(_range.Address, containsBlanks: false);
            return _range;
        }

        /// <summary>
        /// Applies an errors rule.
        /// </summary>
        public ExcelRange Errors() {
            _range.Sheet.AddConditionalErrorsRule(_range.Address, containsErrors: true);
            return _range;
        }

        /// <summary>
        /// Applies a non-errors rule.
        /// </summary>
        public ExcelRange NonErrors() {
            _range.Sheet.AddConditionalErrorsRule(_range.Address, containsErrors: false);
            return _range;
        }

        /// <summary>
        /// Applies a time-period rule.
        /// </summary>
        public ExcelRange TimePeriod(TimePeriodValues timePeriod) {
            _range.Sheet.AddConditionalTimePeriodRule(_range.Address, timePeriod);
            return _range;
        }

        /// <summary>
        /// Applies a time-period rule with a differential fill color.
        /// </summary>
        public ExcelRange TimePeriod(TimePeriodValues timePeriod, string fillColor) {
            _range.Sheet.AddConditionalTimePeriodRule(_range.Address, timePeriod, fillColor: fillColor);
            return _range;
        }

        /// <summary>
        /// Applies a two-color scale.
        /// </summary>
        public ExcelRange ColorScale(OfficeColor startColor, OfficeColor endColor) {
            _range.Sheet.AddConditionalColorScale(_range.Address, startColor, endColor);
            return _range;
        }

        /// <summary>
        /// Applies a two-color scale using hex colors.
        /// </summary>
        public ExcelRange ColorScale(string startColor, string endColor) {
            _range.Sheet.AddConditionalColorScale(_range.Address, startColor, endColor);
            return _range;
        }

        /// <summary>
        /// Applies a data bar.
        /// </summary>
        public ExcelRange DataBar(OfficeColor color) {
            _range.Sheet.AddConditionalDataBar(_range.Address, color);
            return _range;
        }

        /// <summary>
        /// Applies a data bar using a hex color.
        /// </summary>
        public ExcelRange DataBar(string color) {
            _range.Sheet.AddConditionalDataBar(_range.Address, color);
            return _range;
        }

        /// <summary>
        /// Applies an icon set.
        /// </summary>
        public ExcelRange IconSet() {
            return IconSet(IconSetValues.ThreeTrafficLights1);
        }

        /// <summary>
        /// Applies an icon set.
        /// </summary>
        public ExcelRange IconSet(IconSetValues iconSet, bool showValue = true, bool reverseIconOrder = false, double[]? percentThresholds = null, double[]? numberThresholds = null) {
            _range.Sheet.AddConditionalIconSet(_range.Address, iconSet, showValue, reverseIconOrder, percentThresholds, numberThresholds);
            return _range;
        }

        /// <summary>
        /// Removes conditional formatting rules that overlap the range.
        /// </summary>
        public ExcelRange Clear() {
            _range.Sheet.ClearConditionalFormatting(_range.Address);
            return _range;
        }
    }
}
