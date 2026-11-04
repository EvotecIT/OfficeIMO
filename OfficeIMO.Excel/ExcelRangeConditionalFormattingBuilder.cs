using OfficeColor = OfficeIMO.OfficeColor;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Adds conditional formatting rules to an <see cref="ExcelRange"/>. Each method adds one native rule and
    /// returns it attached to the worksheet; edit it with <see cref="ExcelSheet.UpdateConditionalFormattingRule"/>.
    /// </summary>
    /// <remarks>
    /// Existing rules are listed, updated, cloned, reordered and removed through the worksheet
    /// (<see cref="ExcelSheet.GetConditionalFormattingRules(string)"/>, <see cref="ExcelSheet.ClearConditionalFormatting"/>).
    /// <see cref="ExcelSheet.AddConditionalFormattingRule"/> adds a rule from the full rule model, including
    /// multi-area ranges and Office 2010 extension rules.
    /// </remarks>
    public sealed class ExcelRangeConditionalFormattingBuilder {
        private readonly ExcelRange _range;

        internal ExcelRangeConditionalFormattingBuilder(ExcelRange range) {
            _range = range ?? throw new ArgumentNullException(nameof(range));
        }

        /// <summary>Adds a cell-value comparison rule, such as greater than, between or equal.</summary>
        /// <param name="operator">Comparison operator.</param>
        /// <param name="formula1">Primary value or formula.</param>
        /// <param name="formula2">Secondary value or formula for Between and NotBetween.</param>
        /// <param name="fillColor">Optional hex fill applied when the rule matches.</param>
        /// <param name="stopIfTrue">Whether lower-priority rules stop when this rule matches.</param>
        /// <param name="priority">Optional explicit rule priority.</param>
        public ExcelConditionalFormattingInfo CellIs(ExcelConditionalFormattingOperator @operator, string formula1, string? formula2 = null,
            string? fillColor = null, bool stopIfTrue = false, int? priority = null) =>
            _range.Sheet.CreateConditionalCellIsRule(_range.Reference, @operator, formula1, formula2, fillColor, stopIfTrue, priority);

        /// <summary>Adds a formula rule that matches when the formula evaluates to true for the top-left cell.</summary>
        /// <param name="formula">Formula relative to the top-left cell of the range.</param>
        /// <param name="fillColor">Optional hex fill applied when the rule matches.</param>
        /// <param name="stopIfTrue">Whether lower-priority rules stop when this rule matches.</param>
        /// <param name="priority">Optional explicit rule priority.</param>
        public ExcelConditionalFormattingInfo Formula(string formula, string? fillColor = null, bool stopIfTrue = false, int? priority = null) =>
            _range.Sheet.CreateConditionalFormulaRule(_range.Reference, formula, fillColor, stopIfTrue, priority);

        /// <summary>Adds a text rule: contains, does not contain, begins with or ends with.</summary>
        /// <param name="type">ContainsText, NotContainsText, BeginsWith or EndsWith.</param>
        /// <param name="text">Text to match.</param>
        /// <param name="fillColor">Optional hex fill applied when the rule matches.</param>
        /// <param name="stopIfTrue">Whether lower-priority rules stop when this rule matches.</param>
        public ExcelConditionalFormattingInfo Text(ExcelConditionalFormatType type, string text, string? fillColor = null, bool stopIfTrue = false) =>
            _range.Sheet.CreateConditionalTextRule(_range.Reference, type, text, fillColor, stopIfTrue);

        /// <summary>Adds a duplicate-values rule.</summary>
        public ExcelConditionalFormattingInfo DuplicateValues(string? fillColor = null) =>
            _range.Sheet.CreateConditionalDuplicateValuesRule(_range.Reference, fillColor);

        /// <summary>Adds a unique-values rule.</summary>
        public ExcelConditionalFormattingInfo UniqueValues(string? fillColor = null) =>
            _range.Sheet.CreateConditionalUniqueValuesRule(_range.Reference, fillColor);

        /// <summary>Adds an above-average or, with <paramref name="aboveAverage"/> false, a below-average rule.</summary>
        /// <param name="aboveAverage">True for values above the average; false for values below it.</param>
        /// <param name="equalAverage">Whether values equal to the average match.</param>
        /// <param name="standardDeviation">Optional number of standard deviations from the average.</param>
        /// <param name="fillColor">Optional hex fill applied when the rule matches.</param>
        /// <param name="stopIfTrue">Whether lower-priority rules stop when this rule matches.</param>
        public ExcelConditionalFormattingInfo AboveAverage(bool aboveAverage = true, bool equalAverage = false, uint? standardDeviation = null,
            string? fillColor = null, bool stopIfTrue = false) =>
            _range.Sheet.CreateConditionalAboveAverageRule(_range.Reference, aboveAverage, equalAverage, standardDeviation, fillColor, stopIfTrue);

        /// <summary>Adds a top or bottom N (or N percent) rule.</summary>
        /// <param name="rank">Number of items, or percentage when <paramref name="percent"/> is true.</param>
        /// <param name="bottom">True for the lowest values; false for the highest.</param>
        /// <param name="percent">Whether <paramref name="rank"/> is a percentage.</param>
        /// <param name="fillColor">Optional hex fill applied when the rule matches.</param>
        public ExcelConditionalFormattingInfo TopBottom(uint rank, bool bottom = false, bool percent = false, string? fillColor = null) =>
            _range.Sheet.CreateConditionalTopBottomRule(_range.Reference, rank, bottom, percent, fillColor);

        /// <summary>Adds a blanks or, with <paramref name="containsBlanks"/> false, a non-blanks rule.</summary>
        public ExcelConditionalFormattingInfo Blanks(bool containsBlanks = true, bool stopIfTrue = false) =>
            _range.Sheet.CreateConditionalBlanksRule(_range.Reference, containsBlanks, stopIfTrue);

        /// <summary>Adds an errors or, with <paramref name="containsErrors"/> false, a non-errors rule.</summary>
        public ExcelConditionalFormattingInfo Errors(bool containsErrors = true, bool stopIfTrue = false) =>
            _range.Sheet.CreateConditionalErrorsRule(_range.Reference, containsErrors, stopIfTrue);

        /// <summary>Adds a date rule for a time period such as yesterday, this week or last month.</summary>
        public ExcelConditionalFormattingInfo TimePeriod(ExcelConditionalTimePeriod timePeriod, string? fillColor = null, bool stopIfTrue = false) =>
            _range.Sheet.CreateConditionalTimePeriodRule(_range.Reference, timePeriod, fillColor, stopIfTrue);

        /// <summary>Adds a two-color scale from the minimum to the maximum value.</summary>
        public ExcelConditionalFormattingInfo ColorScale(OfficeColor startColor, OfficeColor endColor) =>
            _range.Sheet.CreateConditionalColorScale(_range.Reference, startColor, endColor);

        /// <summary>Adds a two-color scale from hex colors such as F8D7DA.</summary>
        public ExcelConditionalFormattingInfo ColorScale(string startColor, string endColor) =>
            _range.Sheet.CreateConditionalColorScale(_range.Reference, startColor, endColor);

        /// <summary>Adds a data bar.</summary>
        public ExcelConditionalFormattingInfo DataBar(OfficeColor color) =>
            _range.Sheet.CreateConditionalDataBar(_range.Reference, color);

        /// <summary>Adds a data bar from a hex color such as 4472C4.</summary>
        public ExcelConditionalFormattingInfo DataBar(string color) =>
            _range.Sheet.CreateConditionalDataBar(_range.Reference, color);

        /// <summary>Adds an icon set.</summary>
        /// <param name="iconSet">Icon set, such as ThreeTrafficLights1, FourArrows or FiveRating.</param>
        /// <param name="showValue">Whether the cell value is shown next to the icon.</param>
        /// <param name="reverseIconOrder">Whether the icon order is reversed.</param>
        /// <param name="percentThresholds">Optional percent thresholds (0..100), one per icon.</param>
        /// <param name="numberThresholds">Optional number thresholds, one per icon; used instead of percent thresholds.</param>
        public ExcelConditionalFormattingInfo IconSet(ExcelIconSet iconSet = ExcelIconSet.ThreeTrafficLights1, bool showValue = true,
            bool reverseIconOrder = false, double[]? percentThresholds = null, double[]? numberThresholds = null) =>
            _range.Sheet.CreateConditionalIconSet(_range.Reference, iconSet, showValue, reverseIconOrder, percentThresholds, numberThresholds);
    }
}
