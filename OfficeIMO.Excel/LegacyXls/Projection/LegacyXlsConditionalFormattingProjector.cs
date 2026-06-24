using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static class LegacyXlsConditionalFormattingProjector {
        internal static void Project(ExcelSheet sheet, LegacyXlsConditionalFormatting conditionalFormatting) {
            string range = string.Join(" ", conditionalFormatting.Ranges);
            if (string.IsNullOrWhiteSpace(range) || string.IsNullOrWhiteSpace(conditionalFormatting.Formula1)) {
                return;
            }

            uint? differentialFormatId = TryAppendDifferentialFormat(sheet, conditionalFormatting.DifferentialFormat);
            switch (conditionalFormatting.Type) {
                case LegacyXlsConditionalFormattingType.CellIs:
                    if (!conditionalFormatting.Operator.HasValue) {
                        return;
                    }

                    sheet.AddConditionalRule(
                        range,
                        ToOperator(conditionalFormatting.Operator.Value),
                        conditionalFormatting.Formula1,
                        conditionalFormatting.Formula2,
                        conditionalFormatting.StopIfTrue,
                        conditionalFormatting.Priority);
                    ApplyDifferentialFormatId(sheet, differentialFormatId);
                    break;
                case LegacyXlsConditionalFormattingType.Formula:
                    sheet.AddConditionalFormulaRule(range, conditionalFormatting.Formula1, conditionalFormatting.StopIfTrue, conditionalFormatting.Priority);
                    ApplyDifferentialFormatId(sheet, differentialFormatId);
                    break;
            }
        }

        private static void ApplyDifferentialFormatId(ExcelSheet sheet, uint? differentialFormatId) {
            if (differentialFormatId.HasValue) {
                sheet.SetLastConditionalFormattingRuleDifferentialFormatId(differentialFormatId.Value);
            }
        }

        private static uint? TryAppendDifferentialFormat(ExcelSheet sheet, LegacyXlsDifferentialFormat? differentialFormat) {
            if (differentialFormat == null) {
                return null;
            }

            Font? font = TryCreateFont(differentialFormat);
            Fill? fill = TryCreateFill(differentialFormat);
            if (font == null && fill == null) {
                return null;
            }

            var openXmlFormat = new DifferentialFormat();
            if (font != null) {
                openXmlFormat.Append(font);
            }

            if (fill != null) {
                openXmlFormat.Append(fill);
            }

            return sheet.AppendConditionalDifferentialFormat(openXmlFormat);
        }

        private static Font? TryCreateFont(LegacyXlsDifferentialFormat differentialFormat) {
            if (string.IsNullOrWhiteSpace(differentialFormat.FontColor)
                && differentialFormat.FontBold != true
                && differentialFormat.FontItalic != true) {
                return null;
            }

            var font = new Font();
            if (!string.IsNullOrWhiteSpace(differentialFormat.FontColor)) {
                font.Append(new Color { Rgb = differentialFormat.FontColor });
            }

            if (differentialFormat.FontBold == true) {
                font.Append(new Bold());
            }

            if (differentialFormat.FontItalic == true) {
                font.Append(new Italic());
            }

            return font;
        }

        private static Fill? TryCreateFill(LegacyXlsDifferentialFormat differentialFormat) {
            string? color = differentialFormat.FillForegroundColor ?? differentialFormat.FillBackgroundColor;
            if (string.IsNullOrWhiteSpace(color)) {
                return null;
            }

            PatternValues pattern = ToPattern(differentialFormat.FillPattern) ?? PatternValues.Solid;
            var patternFill = new PatternFill {
                PatternType = pattern,
                ForegroundColor = new ForegroundColor { Rgb = color },
                BackgroundColor = new BackgroundColor { Rgb = differentialFormat.FillBackgroundColor ?? color }
            };
            return new Fill(patternFill);
        }

        private static PatternValues? ToPattern(byte? pattern) {
            return pattern switch {
                null => null,
                1 => PatternValues.Solid,
                2 => PatternValues.MediumGray,
                3 => PatternValues.DarkGray,
                4 => PatternValues.LightGray,
                5 => PatternValues.DarkHorizontal,
                6 => PatternValues.DarkVertical,
                7 => PatternValues.DarkDown,
                8 => PatternValues.DarkUp,
                9 => PatternValues.DarkGrid,
                10 => PatternValues.DarkTrellis,
                11 => PatternValues.LightHorizontal,
                12 => PatternValues.LightVertical,
                13 => PatternValues.LightDown,
                14 => PatternValues.LightUp,
                15 => PatternValues.LightGrid,
                16 => PatternValues.LightTrellis,
                17 => PatternValues.Gray125,
                18 => PatternValues.Gray0625,
                _ => null
            };
        }

        private static ConditionalFormattingOperatorValues ToOperator(LegacyXlsConditionalFormattingOperator @operator) {
            return @operator switch {
                LegacyXlsConditionalFormattingOperator.Between => ConditionalFormattingOperatorValues.Between,
                LegacyXlsConditionalFormattingOperator.NotBetween => ConditionalFormattingOperatorValues.NotBetween,
                LegacyXlsConditionalFormattingOperator.Equal => ConditionalFormattingOperatorValues.Equal,
                LegacyXlsConditionalFormattingOperator.NotEqual => ConditionalFormattingOperatorValues.NotEqual,
                LegacyXlsConditionalFormattingOperator.GreaterThan => ConditionalFormattingOperatorValues.GreaterThan,
                LegacyXlsConditionalFormattingOperator.LessThan => ConditionalFormattingOperatorValues.LessThan,
                LegacyXlsConditionalFormattingOperator.GreaterThanOrEqual => ConditionalFormattingOperatorValues.GreaterThanOrEqual,
                _ => ConditionalFormattingOperatorValues.LessThanOrEqual
            };
        }
    }
}
