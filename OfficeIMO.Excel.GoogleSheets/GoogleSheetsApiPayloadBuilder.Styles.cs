using System.Globalization;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static partial class GoogleSheetsApiPayloadBuilder {
        private static GoogleSheetsApiCellFormatPayload? BuildCellFormat(GoogleSheetsCellStyle? style) {
            if (style == null) {
                return null;
            }

            var payload = new GoogleSheetsApiCellFormatPayload {
                NumberFormat = BuildNumberFormat(style),
                BackgroundColor = BuildColor(style.FillColorArgb),
                Borders = BuildBorders(style.Borders),
                HorizontalAlignment = NormalizeHorizontalAlignment(style.HorizontalAlignment),
                VerticalAlignment = NormalizeVerticalAlignment(style.VerticalAlignment),
                WrapStrategy = style.WrapText ? "WRAP" : null,
                Padding = BuildPadding(style.TextIndent),
                TextRotation = BuildTextRotation(style.TextRotation),
            };

            if (style.Bold || style.Italic || style.Underline || style.Strikethrough
                || !string.IsNullOrWhiteSpace(style.FontName) || style.FontSize.HasValue
                || !string.IsNullOrWhiteSpace(style.FontColorArgb)) {
                payload.TextFormat = new GoogleSheetsApiTextFormatPayload {
                    Bold = style.Bold ? true : (bool?)null,
                    Italic = style.Italic ? true : (bool?)null,
                    Underline = style.Underline ? true : (bool?)null,
                    Strikethrough = style.Strikethrough ? true : (bool?)null,
                    FontFamily = style.FontName,
                    FontSize = style.FontSize.HasValue ? Math.Max(1, (int)Math.Round(style.FontSize.Value)) : (int?)null,
                    ForegroundColor = BuildColor(style.FontColorArgb),
                };
            }

            return payload;
        }

        private static GoogleSheetsApiPaddingPayload? BuildPadding(uint? indent) {
            if (!indent.HasValue || indent.Value == 0) {
                return null;
            }

            return new GoogleSheetsApiPaddingPayload {
                Top = 2,
                Right = 2,
                Bottom = 2,
                Left = checked((int)Math.Min(indent.Value * 10U, 250U)),
            };
        }

        private static GoogleSheetsApiTextRotationPayload? BuildTextRotation(int? excelRotation) {
            if (!excelRotation.HasValue) {
                return null;
            }

            if (excelRotation.Value == 255) {
                return new GoogleSheetsApiTextRotationPayload { Vertical = true };
            }

            int angle = excelRotation.Value <= 90
                ? excelRotation.Value
                : 90 - excelRotation.Value;
            return new GoogleSheetsApiTextRotationPayload { Angle = Math.Max(-90, Math.Min(90, angle)) };
        }

        private static GoogleSheetsApiBordersPayload? BuildBorders(GoogleSheetsCellBorders? borders) {
            if (borders == null) {
                return null;
            }

            var payload = new GoogleSheetsApiBordersPayload {
                Left = BuildBorderSide(borders.Left),
                Right = BuildBorderSide(borders.Right),
                Top = BuildBorderSide(borders.Top),
                Bottom = BuildBorderSide(borders.Bottom),
            };

            if (payload.Left == null && payload.Right == null && payload.Top == null && payload.Bottom == null) {
                return null;
            }

            return payload;
        }

        private static GoogleSheetsApiBorderPayload? BuildBorderSide(GoogleSheetsBorderSide? side) {
            if (side == null) {
                return null;
            }

            var style = NormalizeBorderStyle(side.Style);
            var color = BuildColor(side.ColorArgb);
            if (style == null && color == null) {
                return null;
            }

            return new GoogleSheetsApiBorderPayload {
                Style = style ?? "SOLID",
                Color = color,
            };
        }

        private static GoogleSheetsApiNumberFormatPayload? BuildNumberFormat(GoogleSheetsCellStyle style) {
            if (string.IsNullOrWhiteSpace(style.NumberFormatCode) && !style.IsDateLike) {
                return null;
            }

            return new GoogleSheetsApiNumberFormatPayload {
                Type = ResolveNumberFormatType(style),
                Pattern = style.NumberFormatCode,
            };
        }

        private static string ResolveNumberFormatType(GoogleSheetsCellStyle style) {
            if (style.IsDateLike) {
                return "DATE_TIME";
            }

            var pattern = style.NumberFormatCode ?? string.Empty;
            if (pattern.IndexOf('%') >= 0) {
                return "PERCENT";
            }

            if (pattern.IndexOf('$') >= 0 || pattern.IndexOf("z", StringComparison.OrdinalIgnoreCase) >= 0) {
                return "CURRENCY";
            }

            return "NUMBER";
        }

        private static string? NormalizeBorderStyle(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            var normalized = value == null ? string.Empty : value.Trim().ToLowerInvariant();
            return normalized switch {
                "thin" => "SOLID",
                "medium" => "SOLID_MEDIUM",
                "thick" => "SOLID_THICK",
                "double" => "DOUBLE",
                "dashed" => "DASHED",
                "mediumdashed" => "DASHED",
                "dashdot" => "DASHED",
                "mediumdashdot" => "DASHED",
                "dashdotdot" => "DOTTED",
                "mediumdashdotdot" => "DOTTED",
                "dotted" => "DOTTED",
                "hair" => "DOTTED",
                "slantdashdot" => "DASHED",
                _ => "SOLID",
            };
        }

        private static GoogleSheetsApiColorPayload? BuildColor(string? argb) {
            if (string.IsNullOrWhiteSpace(argb) || (argb!.Length != 8 && argb.Length != 6)) {
                return null;
            }

            int offset = argb.Length == 8 ? 2 : 0;
            var red = int.Parse(argb.Substring(offset, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255d;
            var green = int.Parse(argb.Substring(offset + 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255d;
            var blue = int.Parse(argb.Substring(offset + 4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255d;

            return new GoogleSheetsApiColorPayload {
                Red = red,
                Green = green,
                Blue = blue,
            };
        }

        private static GoogleSheetsApiTableRowsPropertiesPayload? BuildTableRowsProperties(GoogleSheetsAddTableRequest table) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            var headerColorStyle = BuildColorStyle(table.HeaderColorArgb);
            var firstBandColorStyle = BuildColorStyle(table.FirstBandColorArgb);
            var secondBandColorStyle = BuildColorStyle(table.SecondBandColorArgb);
            var footerColorStyle = BuildColorStyle(table.FooterColorArgb);

            if (headerColorStyle == null
                && firstBandColorStyle == null
                && secondBandColorStyle == null
                && footerColorStyle == null) {
                return null;
            }

            return new GoogleSheetsApiTableRowsPropertiesPayload {
                HeaderColorStyle = headerColorStyle,
                FirstBandColorStyle = firstBandColorStyle,
                SecondBandColorStyle = secondBandColorStyle,
                FooterColorStyle = footerColorStyle,
            };
        }

        private static GoogleSheetsApiColorStylePayload? BuildColorStyle(string? argb) {
            var color = BuildColor(argb);
            if (color == null) {
                return null;
            }

            return new GoogleSheetsApiColorStylePayload {
                RgbColor = color,
            };
        }

        private static GoogleSheetsApiDataValidationRulePayload? BuildDataValidationRule(GoogleSheetsDataValidationRule? rule) {
            if (rule == null || string.IsNullOrWhiteSpace(rule.ConditionType)) {
                return null;
            }

            return new GoogleSheetsApiDataValidationRulePayload {
                Condition = new GoogleSheetsApiBooleanConditionPayload {
                    Type = rule.ConditionType,
                    Values = rule.Values.Count == 0
                        ? null
                        : rule.Values.Select(value => new GoogleSheetsApiConditionValuePayload {
                            UserEnteredValue = value,
                        }).ToList(),
                },
                Strict = rule.Strict,
                ShowCustomUi = rule.ShowCustomUi,
            };
        }

        private static List<GoogleSheetsApiTextFormatRunPayload>? BuildTextFormatRuns(IReadOnlyList<GoogleSheetsTextFormatRun> runs) {
            if (runs == null || runs.Count == 0) return null;
            return runs.Select(run => new GoogleSheetsApiTextFormatRunPayload {
                StartIndex = run.StartIndex,
                Format = BuildCellFormat(run.Format)?.TextFormat ?? new GoogleSheetsApiTextFormatPayload(),
            }).ToList();
        }

        private static string? NormalizeHorizontalAlignment(string? value) {
            return value switch {
                null => null,
                "" => null,
                "left" => "LEFT",
                "center" => "CENTER",
                "right" => "RIGHT",
                "fill" => "LEFT",
                "justify" => "CENTER",
                _ => value.ToUpperInvariant(),
            };
        }

        private static string? NormalizeVerticalAlignment(string? value) {
            return value switch {
                null => null,
                "" => null,
                "top" => "TOP",
                "center" => "MIDDLE",
                "bottom" => "BOTTOM",
                _ => value.ToUpperInvariant(),
            };
        }

    }
}
