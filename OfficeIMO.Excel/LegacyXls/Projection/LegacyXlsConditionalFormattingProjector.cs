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
            NumberingFormat? numberFormat = TryCreateNumberingFormat(differentialFormat);
            Fill? fill = TryCreateFill(differentialFormat);
            Border? border = TryCreateBorder(differentialFormat);
            if (font == null && numberFormat == null && fill == null && border == null) {
                return null;
            }

            var openXmlFormat = new DifferentialFormat();
            if (font != null) {
                openXmlFormat.Append(font);
            }

            if (numberFormat != null) {
                openXmlFormat.Append(numberFormat);
            }

            if (fill != null) {
                openXmlFormat.Append(fill);
            }

            if (border != null) {
                openXmlFormat.Append(border);
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

        private static NumberingFormat? TryCreateNumberingFormat(LegacyXlsDifferentialFormat differentialFormat) {
            if (string.IsNullOrWhiteSpace(differentialFormat.NumberFormatCode)) {
                return null;
            }

            return new NumberingFormat {
                NumberFormatId = differentialFormat.NumberFormatId ?? 164U,
                FormatCode = differentialFormat.NumberFormatCode
            };
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

        private static Border? TryCreateBorder(LegacyXlsDifferentialFormat differentialFormat) {
            LegacyXlsDifferentialBorder? border = differentialFormat.Border;
            if (border?.HasAnySide != true) {
                return null;
            }

            var openXmlBorder = new Border();
            if (border.Left != null) {
                openXmlBorder.LeftBorder = CreateBorderSide<LeftBorder>(border.Left);
            }

            if (border.Right != null) {
                openXmlBorder.RightBorder = CreateBorderSide<RightBorder>(border.Right);
            }

            if (border.Top != null) {
                openXmlBorder.TopBorder = CreateBorderSide<TopBorder>(border.Top);
            }

            if (border.Bottom != null) {
                openXmlBorder.BottomBorder = CreateBorderSide<BottomBorder>(border.Bottom);
            }

            return openXmlBorder;
        }

        private static T CreateBorderSide<T>(LegacyXlsDifferentialBorderSide side) where T : BorderPropertiesType, new() {
            var openXmlSide = new T();
            BorderStyleValues? style = ToBorderStyle(side.Style);
            if (style.HasValue) {
                openXmlSide.Style = style.Value;
            }

            if (!string.IsNullOrWhiteSpace(side.Color)) {
                openXmlSide.Color = new Color { Rgb = side.Color };
            }

            return openXmlSide;
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

        private static BorderStyleValues? ToBorderStyle(ushort style) {
            return style switch {
                1 => BorderStyleValues.Thin,
                2 => BorderStyleValues.Medium,
                3 => BorderStyleValues.Dashed,
                4 => BorderStyleValues.Dotted,
                5 => BorderStyleValues.Thick,
                6 => BorderStyleValues.Double,
                7 => BorderStyleValues.Hair,
                8 => BorderStyleValues.MediumDashed,
                9 => BorderStyleValues.DashDot,
                10 => BorderStyleValues.MediumDashDot,
                11 => BorderStyleValues.DashDotDot,
                12 => BorderStyleValues.MediumDashDotDot,
                13 => BorderStyleValues.SlantDashDot,
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
