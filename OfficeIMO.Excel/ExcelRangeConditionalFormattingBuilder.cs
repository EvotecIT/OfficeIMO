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
        public ExcelRange CellIs(ConditionalFormattingOperatorValues @operator, string formula1, string? formula2 = null) {
            _range.Sheet.AddConditionalRule(_range.Address, @operator, formula1, formula2);
            return _range;
        }

        /// <summary>
        /// Applies a greater-than cell-value rule.
        /// </summary>
        public ExcelRange GreaterThan(string formula) {
            return CellIs(ConditionalFormattingOperatorValues.GreaterThan, formula);
        }

        /// <summary>
        /// Applies a less-than cell-value rule.
        /// </summary>
        public ExcelRange LessThan(string formula) {
            return CellIs(ConditionalFormattingOperatorValues.LessThan, formula);
        }

        /// <summary>
        /// Applies a between cell-value rule.
        /// </summary>
        public ExcelRange Between(string formula1, string formula2) {
            return CellIs(ConditionalFormattingOperatorValues.Between, formula1, formula2);
        }

        /// <summary>
        /// Applies a formula-based rule.
        /// </summary>
        public ExcelRange Formula(string formula, bool stopIfTrue = false) {
            _range.Sheet.AddConditionalFormulaRule(_range.Address, formula, stopIfTrue);
            return _range;
        }

        /// <summary>
        /// Applies a duplicate-values rule.
        /// </summary>
        public ExcelRange DuplicateValues() {
            _range.Sheet.AddConditionalDuplicateValuesRule(_range.Address);
            return _range;
        }

        /// <summary>
        /// Applies a top-values rule.
        /// </summary>
        public ExcelRange Top(uint rank, bool percent = false) {
            _range.Sheet.AddConditionalTopBottomRule(_range.Address, rank, bottom: false, percent: percent);
            return _range;
        }

        /// <summary>
        /// Applies a bottom-values rule.
        /// </summary>
        public ExcelRange Bottom(uint rank, bool percent = false) {
            _range.Sheet.AddConditionalTopBottomRule(_range.Address, rank, bottom: true, percent: percent);
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
