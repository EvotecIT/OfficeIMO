using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SixLabors.Fonts;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using SixLaborsColor = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Adds a conditional formatting rule to the specified range.
        /// </summary>
        /// <param name="range">A1-style range to apply the rule to.</param>
        /// <param name="operator">Comparison operator for the rule.</param>
        /// <param name="formula1">Primary formula or value.</param>
        /// <param name="formula2">Optional secondary formula or value.</param>
        public void AddConditionalRule(string range, ConditionalFormattingOperatorValues @operator, string formula1, string? formula2 = null) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            WriteLock(() => {
                Worksheet worksheet = _worksheetPart.Worksheet;

                int priority = 1;
                var existingRules = worksheet.Descendants<ConditionalFormattingRule>();
                if (existingRules.Any()) {
                    priority = existingRules.Max(r => r.Priority?.Value ?? 0) + 1;
                }

                ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };

                ConditionalFormattingRule rule = new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.CellIs,
                    Operator = @operator,
                    Priority = priority
                };

                rule.Append(new Formula(formula1));
                if (formula2 != null) {
                    rule.Append(new Formula(formula2));
                }

                conditionalFormatting.Append(rule);
                
                // Insert ConditionalFormatting after AutoFilter but before TableParts
                var autoFilter = worksheet.Elements<AutoFilter>().FirstOrDefault();
                var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
                
                if (tableParts != null) {
                    worksheet.InsertBefore(conditionalFormatting, tableParts);
                } else if (autoFilter != null) {
                    worksheet.InsertAfter(conditionalFormatting, autoFilter);
                } else {
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        worksheet.InsertAfter(conditionalFormatting, sheetData);
                    } else {
                        worksheet.Append(conditionalFormatting);
                    }
                }
                
                worksheet.Save();
            });
        }

        private static string ConvertColor(SixLaborsColor color) {
            return "FF" + color.ToHexColor();
        }

        /// <summary>
        /// Adds a two-color scale conditional format to a range.
        /// </summary>
        /// <param name="range">A1-style range to format.</param>
        /// <param name="startColor">Starting color of the scale.</param>
        /// <param name="endColor">Ending color of the scale.</param>
        public void AddConditionalColorScale(string range, SixLaborsColor startColor, SixLaborsColor endColor) {
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

            WriteLock(() => {
                Worksheet worksheet = _worksheetPart.Worksheet;

                int priority = 1;
                var existingRules = worksheet.Descendants<ConditionalFormattingRule>();
                if (existingRules.Any()) {
                    priority = existingRules.Max(r => r.Priority?.Value ?? 0) + 1;
                }

                ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };

                ConditionalFormattingRule rule = new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.ColorScale,
                    Priority = priority
                };

                ColorScale colorScale = new ColorScale();
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
                colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = startColor });
                colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = endColor });
                rule.Append(colorScale);

                conditionalFormatting.Append(rule);
                
                // Insert ConditionalFormatting after AutoFilter but before TableParts
                var autoFilter = worksheet.Elements<AutoFilter>().FirstOrDefault();
                var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
                
                if (tableParts != null) {
                    worksheet.InsertBefore(conditionalFormatting, tableParts);
                } else if (autoFilter != null) {
                    worksheet.InsertAfter(conditionalFormatting, autoFilter);
                } else {
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        worksheet.InsertAfter(conditionalFormatting, sheetData);
                    } else {
                        worksheet.Append(conditionalFormatting);
                    }
                }
                
                worksheet.Save();
            });
        }

        /// <summary>
        /// Adds a data bar conditional format to a range.
        /// </summary>
        /// <param name="range">A1-style range to format.</param>
        /// <param name="color">Bar color.</param>
        public void AddConditionalDataBar(string range, SixLaborsColor color) {
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

            WriteLock(() => {
                Worksheet worksheet = _worksheetPart.Worksheet;

                int priority = 1;
                var existingRules = worksheet.Descendants<ConditionalFormattingRule>();
                if (existingRules.Any()) {
                    priority = existingRules.Max(r => r.Priority?.Value ?? 0) + 1;
                }

                ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };

                ConditionalFormattingRule rule = new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.DataBar,
                    Priority = priority
                };

                DataBar dataBar = new DataBar { ShowValue = true };
                dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
                dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
                dataBar.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = color });
                rule.Append(dataBar);

                conditionalFormatting.Append(rule);
                
                // Insert ConditionalFormatting after AutoFilter but before TableParts
                var autoFilter = worksheet.Elements<AutoFilter>().FirstOrDefault();
                var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
                
                if (tableParts != null) {
                    worksheet.InsertBefore(conditionalFormatting, tableParts);
                } else if (autoFilter != null) {
                    worksheet.InsertAfter(conditionalFormatting, autoFilter);
                } else {
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        worksheet.InsertAfter(conditionalFormatting, sheetData);
                    } else {
                        worksheet.Append(conditionalFormatting);
                    }
                }
                
                worksheet.Save();
            });
        }

        /// <summary>
        /// Adds an icon set conditional format to a range.
        /// </summary>
        /// <param name="range">A1-style range to format.</param>
        /// <param name="iconSet">Icon set type (e.g., ThreeTrafficLights1, ThreeSymbols, FourArrows, FiveRatings).</param>
        /// <param name="showValue">Whether to display the underlying cell values.</param>
        /// <param name="reverseIconOrder">Reverse icon order.</param>
        public void AddConditionalIconSet(string range, IconSetValues iconSet, bool showValue, bool reverseIconOrder)
        {
            AddConditionalIconSet(range, iconSet, showValue, reverseIconOrder, percentThresholds: null, numberThresholds: null);
        }

        /// <summary>
        /// Adds an icon set conditional format to a range with optional explicit thresholds.
        /// Provide either <paramref name="percentThresholds"/> (0..100) or <paramref name="numberThresholds"/> as absolute values.
        /// The number of thresholds must match the icon count for the selected icon set (3/4/5).
        /// </summary>
        public void AddConditionalIconSet(string range, IconSetValues iconSet, bool showValue, bool reverseIconOrder, double[]? percentThresholds, double[]? numberThresholds)
        {
            if (string.IsNullOrEmpty(range)) throw new ArgumentNullException(nameof(range));

            WriteLock(() =>
            {
                Worksheet worksheet = _worksheetPart.Worksheet;

                int priority = 1;
                var existingRules = worksheet.Descendants<ConditionalFormattingRule>();
                if (existingRules.Any()) {
                    priority = existingRules.Max(r => r.Priority?.Value ?? 0) + 1;
                }

                ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };

                ConditionalFormattingRule rule = new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.IconSet,
                    Priority = priority
                };

                var icon = new IconSet { IconSetValue = iconSet, ShowValue = showValue, Reverse = reverseIconOrder };
                // Schema requires cfvo count to match icon count.
                int count;
                var setName = iconSet.ToString();
                if (setName.StartsWith("Three", System.StringComparison.Ordinal)) count = 3;
                else if (setName.StartsWith("Four", System.StringComparison.Ordinal)) count = 4;
                else count = 5;

                if (numberThresholds != null && numberThresholds.Length == count)
                {
                    for (int i = 0; i < count; i++)
                    {
                        var cfvo = new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Number };
                        cfvo.Val = numberThresholds[i].ToString(System.Globalization.CultureInfo.InvariantCulture);
                        icon.Append(cfvo);
                    }
                }
                else if (percentThresholds != null && percentThresholds.Length == count)
                {
                    for (int i = 0; i < count; i++)
                    {
                        var cfvo = new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percent };
                        cfvo.Val = percentThresholds[i].ToString(System.Globalization.CultureInfo.InvariantCulture);
                        icon.Append(cfvo);
                    }
                }
                else
                {
                    // Defaults: spread evenly across percent bands
                    int[] perc = count == 3 ? new[] { 0, 33, 67 } : count == 4 ? new[] { 0, 25, 50, 75 } : new[] { 0, 20, 40, 60, 80 };
                    for (int i = 0; i < perc.Length; i++)
                    {
                        var cfvo = new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percent };
                        cfvo.Val = perc[i].ToString(System.Globalization.CultureInfo.InvariantCulture);
                        icon.Append(cfvo);
                    }
                }
                rule.Append(icon);
                conditionalFormatting.Append(rule);

                // Insert ConditionalFormatting after AutoFilter but before TableParts
                var autoFilter = worksheet.Elements<AutoFilter>().FirstOrDefault();
                var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();

                if (tableParts != null) {
                    worksheet.InsertBefore(conditionalFormatting, tableParts);
                } else if (autoFilter != null) {
                    worksheet.InsertAfter(conditionalFormatting, autoFilter);
                } else {
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        worksheet.InsertAfter(conditionalFormatting, sheetData);
                    } else {
                        worksheet.Append(conditionalFormatting);
                    }
                }

                worksheet.Save();
            });
        }

        /// <summary>
        /// Overload with common defaults for convenience.
        /// </summary>
        public void AddConditionalIconSet(string range)
            => AddConditionalIconSet(range, IconSetValues.ThreeTrafficLights1, showValue: true, reverseIconOrder: false);

    }
}

