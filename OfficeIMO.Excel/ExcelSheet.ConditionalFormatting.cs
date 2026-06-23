using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeColor = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Adds a conditional formatting rule to the specified range.
        /// </summary>
        /// <param name="range">A1-style range to apply the rule to.</param>
        /// <param name="operator">Comparison operator for the rule.</param>
        /// <param name="formula1">Primary formula or value.</param>
        /// <param name="formula2">Optional secondary formula or value.</param>
        /// <param name="stopIfTrue">Whether lower-priority rules should stop when this rule evaluates to true.</param>
        /// <param name="priority">Optional explicit rule priority.</param>
        public void AddConditionalRule(string range, ConditionalFormattingOperatorValues @operator, string formula1, string? formula2 = null, bool stopIfTrue = false, int? priority = null) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            WriteLockWorksheetPreparationOnly(() => {
                Worksheet worksheet = WorksheetRoot;

                ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };

                ConditionalFormattingRule rule = new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.CellIs,
                    Operator = @operator,
                    Priority = priority ?? GetNextConditionalFormattingPriority(),
                    StopIfTrue = stopIfTrue
                };

                rule.Append(new Formula(formula1));
                if (formula2 != null) {
                    rule.Append(new Formula(formula2));
                }

                conditionalFormatting.Append(rule);
                InsertConditionalFormatting(conditionalFormatting);

            });
        }

        private static string ConvertColor(OfficeColor color) {
            return "FF" + color.ToHexColor();
        }

        /// <summary>
        /// Adds a two-color scale conditional format to a range.
        /// </summary>
        /// <param name="range">A1-style range to format.</param>
        /// <param name="startColor">Starting color of the scale.</param>
        /// <param name="endColor">Ending color of the scale.</param>
        public void AddConditionalColorScale(string range, OfficeColor startColor, OfficeColor endColor) {
            AddConditionalColorScale(range, ConvertColor(startColor), ConvertColor(endColor));
        }

        /// <summary>
        /// Adds a two-color scale conditional format to a range using hex colors.
        /// </summary>
        /// <param name="range">A1-style range to format.</param>
        /// <param name="startColor">Starting color in hex (e.g. FF0000).</param>
        /// <param name="endColor">Ending color in hex.</param>
        public void AddConditionalColorScale(string range, string startColor, string endColor) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            WriteLockWorksheetPreparationOnly(() => {
                Worksheet worksheet = WorksheetRoot;

                ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };

                ConditionalFormattingRule rule = new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.ColorScale,
                    Priority = GetNextConditionalFormattingPriority()
                };

                ColorScale colorScale = new ColorScale();
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
                colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = startColor });
                colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = endColor });
                rule.Append(colorScale);

                conditionalFormatting.Append(rule);
                InsertConditionalFormatting(conditionalFormatting);

            });
        }

        /// <summary>
        /// Adds a data bar conditional format to a range.
        /// </summary>
        /// <param name="range">A1-style range to format.</param>
        /// <param name="color">Bar color.</param>
        public void AddConditionalDataBar(string range, OfficeColor color) {
            AddConditionalDataBar(range, ConvertColor(color));
        }

        /// <summary>
        /// Adds a data bar conditional format to a range using a hex color.
        /// </summary>
        /// <param name="range">A1-style range to format.</param>
        /// <param name="color">Bar color in hex (e.g. FF0000).</param>
        public void AddConditionalDataBar(string range, string color) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            WriteLockWorksheetPreparationOnly(() => {
                Worksheet worksheet = WorksheetRoot;

                ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };

                ConditionalFormattingRule rule = new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.DataBar,
                    Priority = GetNextConditionalFormattingPriority()
                };

                DataBar dataBar = new DataBar { ShowValue = true };
                dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
                dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
                dataBar.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = color });
                rule.Append(dataBar);

                conditionalFormatting.Append(rule);
                InsertConditionalFormatting(conditionalFormatting);

            });
        }

        /// <summary>
        /// Adds an icon set conditional format to a range.
        /// </summary>
        /// <param name="range">A1-style range to format.</param>
        /// <param name="iconSet">Icon set type (e.g., ThreeTrafficLights1, ThreeSymbols, FourArrows, FiveRatings).</param>
        /// <param name="showValue">Whether to display the underlying cell values.</param>
        /// <param name="reverseIconOrder">Reverse icon order.</param>
        public void AddConditionalIconSet(string range, IconSetValues iconSet, bool showValue, bool reverseIconOrder) {
            AddConditionalIconSet(range, iconSet, showValue, reverseIconOrder, percentThresholds: null, numberThresholds: null);
        }

        /// <summary>
        /// Adds an icon set conditional format to a range with optional explicit thresholds.
        /// Provide either <paramref name="percentThresholds"/> (0..100) or <paramref name="numberThresholds"/> as absolute values.
        /// The number of thresholds must match the icon count for the selected icon set (3/4/5).
        /// </summary>
        public void AddConditionalIconSet(string range, IconSetValues iconSet, bool showValue, bool reverseIconOrder, double[]? percentThresholds, double[]? numberThresholds) {
            if (string.IsNullOrEmpty(range)) throw new ArgumentNullException(nameof(range));

            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            WriteLockWorksheetPreparationOnly(() => {
                Worksheet worksheet = WorksheetRoot;

                ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };

                ConditionalFormattingRule rule = new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.IconSet,
                    Priority = GetNextConditionalFormattingPriority()
                };

                var icon = new IconSet { IconSetValue = iconSet, ShowValue = showValue, Reverse = reverseIconOrder };
                // Schema requires cfvo count to match icon count.
                int count;
                var setName = iconSet.ToString();
                if (setName.StartsWith("Three", System.StringComparison.Ordinal)) count = 3;
                else if (setName.StartsWith("Four", System.StringComparison.Ordinal)) count = 4;
                else count = 5;

                if (numberThresholds != null && numberThresholds.Length == count) {
                    for (int i = 0; i < count; i++) {
                        var cfvo = new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Number };
                        cfvo.Val = numberThresholds[i].ToString(System.Globalization.CultureInfo.InvariantCulture);
                        icon.Append(cfvo);
                    }
                } else if (percentThresholds != null && percentThresholds.Length == count) {
                    for (int i = 0; i < count; i++) {
                        var cfvo = new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percent };
                        cfvo.Val = percentThresholds[i].ToString(System.Globalization.CultureInfo.InvariantCulture);
                        icon.Append(cfvo);
                    }
                } else {
                    // Defaults: spread evenly across percent bands
                    int[] perc = count == 3 ? new[] { 0, 33, 67 } : count == 4 ? new[] { 0, 25, 50, 75 } : new[] { 0, 20, 40, 60, 80 };
                    for (int i = 0; i < perc.Length; i++) {
                        var cfvo = new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percent };
                        cfvo.Val = perc[i].ToString(System.Globalization.CultureInfo.InvariantCulture);
                        icon.Append(cfvo);
                    }
                }
                rule.Append(icon);
                conditionalFormatting.Append(rule);
                InsertConditionalFormatting(conditionalFormatting);

            });
        }

        /// <summary>
        /// Overload with common defaults for convenience.
        /// </summary>
        public void AddConditionalIconSet(string range)
            => AddConditionalIconSet(range, IconSetValues.ThreeTrafficLights1, showValue: true, reverseIconOrder: false);

    }
}

