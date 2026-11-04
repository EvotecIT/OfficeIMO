using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeColor = OfficeIMO.OfficeColor;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Adds a cell-value comparison rule; authored through <see cref="ExcelRangeConditionalFormattingBuilder.CellIs"/>.</summary>
        internal ExcelConditionalFormattingInfo CreateConditionalCellIsRule(string range, ExcelConditionalFormattingOperator @operator, string formula1, string? formula2, OfficeColor? fillColor, bool stopIfTrue, int? priority) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            ExcelConditionalFormattingInfo? createdRule = null;
            WriteLockWorksheetPreparationOnly(() => {
                _excelDocument.EnsureWorkbookThemeAndStyles();
                Worksheet worksheet = WorksheetRoot;

                ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };

                ConditionalFormattingRule rule = new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.CellIs,
                    Operator = @operator.ToOpenXml(),
                    Priority = priority ?? GetNextConditionalFormattingPriority(),
                    StopIfTrue = stopIfTrue
                };

                rule.Append(new Formula(formula1));
                if (formula2 != null) {
                    rule.Append(new Formula(formula2));
                }

                if (fillColor.HasValue) {
                    rule.FormatId = GetOrCreateDifferentialFillFormatId(fillColor.Value);
                }

                conditionalFormatting.Append(rule);
                InsertConditionalFormatting(conditionalFormatting);
                createdRule = ReadConditionalFormattingInfo(rule, range,
                    _excelDocument.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet, _excelDocument.WorkbookPartRoot);

            });
            return createdRule!;
        }

        internal ExcelConditionalFormattingInfo CreateConditionalColorScale(string range, OfficeColor startColor, OfficeColor endColor) {
            return CreateConditionalColorScaleCore(range, ExcelColor.ToSpreadsheetColor(startColor, _excelDocument.EffectiveCreationTheme), ExcelColor.ToSpreadsheetColor(endColor, _excelDocument.EffectiveCreationTheme));
        }

        internal ExcelConditionalFormattingInfo CreateConditionalColorScale(string range, OfficeColor startColor, OfficeColor midColor, OfficeColor endColor) =>
            CreateConditionalColorScaleCore(range,
                ExcelColor.ToSpreadsheetColor(startColor, _excelDocument.EffectiveCreationTheme),
                ExcelColor.ToSpreadsheetColor(midColor, _excelDocument.EffectiveCreationTheme),
                ExcelColor.ToSpreadsheetColor(endColor, _excelDocument.EffectiveCreationTheme));

        /// <summary>Adds the native color-scale rule through the existing worksheet mutation path.</summary>
        private ExcelConditionalFormattingInfo CreateConditionalColorScaleCore(string range,
            params DocumentFormat.OpenXml.Spreadsheet.Color[] colors) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            ExcelConditionalFormattingInfo? createdRule = null;
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
                if (colors.Length == 3) colorScale.Append(new ConditionalFormatValueObject {
                    Type = ConditionalFormatValueObjectValues.Percentile, Val = "50"
                });
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
                foreach (var color in colors) colorScale.Append(color);
                rule.Append(colorScale);

                conditionalFormatting.Append(rule);
                InsertConditionalFormatting(conditionalFormatting);
                createdRule = ReadConditionalFormattingInfo(rule, range,
                    _excelDocument.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet, _excelDocument.WorkbookPartRoot);

            });
            return createdRule!;
        }

        internal ExcelConditionalFormattingInfo CreateConditionalDataBar(string range, OfficeColor color) {
            return CreateConditionalDataBarCore(range, ExcelColor.ToSpreadsheetColor(color, _excelDocument.EffectiveCreationTheme));
        }

        /// <summary>Adds the native data-bar rule through the existing worksheet mutation path.</summary>
        private ExcelConditionalFormattingInfo CreateConditionalDataBarCore(string range,
            DocumentFormat.OpenXml.Spreadsheet.Color color) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            ExcelConditionalFormattingInfo? createdRule = null;
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
                dataBar.Append(color);
                rule.Append(dataBar);

                conditionalFormatting.Append(rule);
                InsertConditionalFormatting(conditionalFormatting);
                createdRule = ReadConditionalFormattingInfo(rule, range,
                    _excelDocument.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet, _excelDocument.WorkbookPartRoot);

            });
            return createdRule!;
        }

        /// <summary>
        /// Adds an icon set with optional explicit thresholds. Provide either percent thresholds (0..100) or number
        /// thresholds; the count must match the icon count for the selected icon set (3/4/5).
        /// </summary>
        internal ExcelConditionalFormattingInfo CreateConditionalIconSet(string range, ExcelIconSet iconSet, bool showValue, bool reverseIconOrder, double[]? percentThresholds, double[]? numberThresholds) {
            if (string.IsNullOrEmpty(range)) throw new ArgumentNullException(nameof(range));

            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            ExcelConditionalFormattingInfo? createdRule = null;
            WriteLockWorksheetPreparationOnly(() => {
                Worksheet worksheet = WorksheetRoot;

                ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };

                ConditionalFormattingRule rule = new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.IconSet,
                    Priority = GetNextConditionalFormattingPriority()
                };

                var icon = new IconSet { IconSetValue = iconSet.ToOpenXml(), ShowValue = showValue, Reverse = reverseIconOrder };
                int count = ResolveIconSetThresholdCount(iconSet);

                if (numberThresholds != null && numberThresholds.Length == count) {
                    for (int i = 0; i < count; i++) {
                        var cfvo = new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Number };
                        cfvo.Val = InvariantNumberText.Get(numberThresholds[i]);
                        icon.Append(cfvo);
                    }
                } else if (percentThresholds != null && percentThresholds.Length == count) {
                    for (int i = 0; i < count; i++) {
                        var cfvo = new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percent };
                        cfvo.Val = InvariantNumberText.Get(percentThresholds[i]);
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
                createdRule = ReadConditionalFormattingInfo(rule, range,
                    _excelDocument.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet, _excelDocument.WorkbookPartRoot);

            });
            return createdRule!;
        }

        private static int ResolveIconSetThresholdCount(ExcelIconSet iconSet) {
            if (iconSet == ExcelIconSet.ThreeArrows ||
                iconSet == ExcelIconSet.ThreeArrowsGray ||
                iconSet == ExcelIconSet.ThreeFlags ||
                iconSet == ExcelIconSet.ThreeSigns ||
                iconSet == ExcelIconSet.ThreeSymbols ||
                iconSet == ExcelIconSet.ThreeSymbols2 ||
                iconSet == ExcelIconSet.ThreeTrafficLights1 ||
                iconSet == ExcelIconSet.ThreeTrafficLights2) {
                return 3;
            }

            if (iconSet == ExcelIconSet.FourArrows ||
                iconSet == ExcelIconSet.FourArrowsGray ||
                iconSet == ExcelIconSet.FourRating ||
                iconSet == ExcelIconSet.FourRedToBlack ||
                iconSet == ExcelIconSet.FourTrafficLights) {
                return 4;
            }

            return 5;
        }

    }
}
