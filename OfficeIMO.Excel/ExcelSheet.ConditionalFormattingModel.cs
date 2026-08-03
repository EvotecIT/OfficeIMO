using System.Globalization;
using System.Xml;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static void ValidateConditionalFormattingDefinition(
            ExcelConditionalFormattingInfo definition,
            bool validateFormulas,
            bool validateVisual,
            bool validateStyle,
            bool allowUnknownType,
            bool updatingExisting) {
            IReadOnlyList<string> formulas = definition.Formulas ?? Array.Empty<string>();
            IReadOnlyList<string> colorScaleColors = definition.ColorScaleColors ?? Array.Empty<string>();
            IReadOnlyList<ExcelConditionalFormatThreshold> colorScaleThresholds = definition.ColorScaleThresholds ?? Array.Empty<ExcelConditionalFormatThreshold>();
            IReadOnlyList<ExcelConditionalFormatThreshold> dataBarThresholds = definition.DataBarThresholds ?? Array.Empty<ExcelConditionalFormatThreshold>();
            IReadOnlyList<ExcelConditionalIconSetThreshold> iconSetThresholds = definition.IconSetThresholds ?? Array.Empty<ExcelConditionalIconSetThreshold>();
            IReadOnlyList<ExcelConditionalFormatIcon> customIcons = definition.CustomIcons ?? Array.Empty<ExcelConditionalFormatIcon>();
            bool preservingExistingVisual = updatingExisting &&
                string.Equals(definition.ProjectedType, definition.Type, StringComparison.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(definition.Type)) {
                throw new ArgumentException("A conditional-formatting rule type is required.", nameof(definition));
            }
            if (!allowUnknownType && !IsKnownConditionalFormattingType(definition.Type)) {
                throw new ArgumentException($"Conditional-formatting rule type '{definition.Type}' is not supported for authoring.", nameof(definition));
            }
            if (definition.Priority < 0) {
                throw new ArgumentOutOfRangeException(nameof(definition), "Conditional-formatting priority cannot be negative.");
            }
            if (!string.IsNullOrEmpty(definition.Text)) XmlConvert.VerifyXmlChars(definition.Text);
            if (!string.IsNullOrWhiteSpace(definition.Operator) && !IsKnownConditionalFormattingOperator(definition.Operator)) {
                throw new ArgumentException($"Conditional-formatting operator '{definition.Operator}' is not supported.", nameof(definition));
            }
            if (!string.IsNullOrWhiteSpace(definition.TimePeriod) && !IsKnownConditionalFormattingTimePeriod(definition.TimePeriod)) {
                throw new ArgumentException($"Conditional-formatting time period '{definition.TimePeriod}' is not supported.", nameof(definition));
            }
            if (validateStyle) {
                if (!string.IsNullOrWhiteSpace(definition.DifferentialFillColorArgb)) NormalizeHexColor(definition.DifferentialFillColorArgb!);
                if (!string.IsNullOrWhiteSpace(definition.DifferentialFontColorArgb)) NormalizeHexColor(definition.DifferentialFontColorArgb!);
                if (!string.IsNullOrWhiteSpace(definition.DifferentialFontName)) XmlConvert.VerifyXmlChars(definition.DifferentialFontName!);
                ValidateConditionalFormattingBorder(definition.DifferentialBorder);
                if (definition.DifferentialFontSize is <= 0D || double.IsNaN(definition.DifferentialFontSize ?? 1D) || double.IsInfinity(definition.DifferentialFontSize ?? 1D)) {
                    throw new ArgumentOutOfRangeException(nameof(definition), "Conditional-formatting font size must be finite and greater than zero.");
                }
            }

            if (validateFormulas) {
                foreach (string formula in formulas) {
                    if (string.IsNullOrWhiteSpace(formula)) {
                        throw new ArgumentException("Conditional-formatting formulas cannot be empty.", nameof(definition));
                    }
                    XmlConvert.VerifyXmlChars(formula);
                }
                if (string.Equals(definition.Type, "Expression", StringComparison.OrdinalIgnoreCase) && formulas.Count == 0) {
                    throw new ArgumentException("Expression conditional-formatting rules require a formula.", nameof(definition));
                }
                if (string.Equals(definition.Type, "CellIs", StringComparison.OrdinalIgnoreCase) &&
                    (!IsKnownConditionalFormattingOperator(definition.Operator) ||
                     (IsBetweenConditionalFormattingOperator(definition.Operator) ? formulas.Count != 2 : formulas.Count != 1))) {
                    throw new ArgumentException("Cell-is conditional-formatting rules require a supported operator and one formula, or two formulas for between/not-between.", nameof(definition));
                }
            }

            if (string.Equals(definition.Type, "Top10", StringComparison.OrdinalIgnoreCase) && definition.TopBottomRank is null or 0U) {
                throw new ArgumentException("Top/bottom conditional-formatting rules require a positive rank.", nameof(definition));
            }
            if (IsTextConditionalFormattingType(definition.Type) && string.IsNullOrEmpty(definition.Text)) {
                throw new ArgumentException("Text conditional-formatting rules require text.", nameof(definition));
            }
            if (string.Equals(definition.Type, "TimePeriod", StringComparison.OrdinalIgnoreCase) &&
                !IsKnownConditionalFormattingTimePeriod(definition.TimePeriod)) {
                throw new ArgumentException("Time-period conditional-formatting rules require a supported time period.", nameof(definition));
            }
            if (definition.AboveAverageStdDev.HasValue && definition.AboveAverageStdDev.Value < 0) {
                throw new ArgumentOutOfRangeException(nameof(definition), "Above-average standard deviation cannot be negative.");
            }

            if (!validateVisual) return;
            if (string.Equals(definition.Type, "ColorScale", StringComparison.OrdinalIgnoreCase)) {
                if (colorScaleThresholds.Count < 2 || colorScaleThresholds.Count > 3 ||
                    colorScaleColors.Count != colorScaleThresholds.Count && !(preservingExistingVisual && colorScaleColors.Count == 0) ||
                    !preservingExistingVisual && colorScaleColors.Any(string.IsNullOrWhiteSpace)) {
                    throw new ArgumentException("Color-scale rules require matching sets of two or three thresholds and colors.", nameof(definition));
                }
                ValidateConditionalFormattingThresholds(colorScaleThresholds.Select(item => (item.Type, item.Value)), definition.Source);
                ValidateConditionalFormattingColors(colorScaleColors);
            } else if (string.Equals(definition.Type, "DataBar", StringComparison.OrdinalIgnoreCase)) {
                if (dataBarThresholds.Count != 2 || string.IsNullOrWhiteSpace(definition.DataBarColor) && !preservingExistingVisual) {
                    throw new ArgumentException("Data-bar rules require exactly two thresholds and a fill color.", nameof(definition));
                }
                ValidateConditionalFormattingThresholds(dataBarThresholds.Select(item => (item.Type, item.Value)), definition.Source);
                if (definition.DataBarMinimumLength > 100U || definition.DataBarMaximumLength > 100U ||
                    definition.DataBarMinimumLength.HasValue && definition.DataBarMaximumLength.HasValue &&
                    definition.DataBarMinimumLength.Value > definition.DataBarMaximumLength.Value) {
                    throw new ArgumentOutOfRangeException(nameof(definition), "Data-bar lengths must be between 0 and 100 with minimum not greater than maximum.");
                }
                ValidateConditionalFormattingColors(new[] {
                    definition.DataBarColor,
                    definition.DataBarBorderColor,
                    definition.DataBarNegativeColor,
                    definition.DataBarNegativeBorderColor,
                    definition.DataBarAxisColor
                }.Where(color => !string.IsNullOrWhiteSpace(color)).Select(color => color!));
            } else if (string.Equals(definition.Type, "IconSet", StringComparison.OrdinalIgnoreCase)) {
                int count = iconSetThresholds.Count;
                string? iconSet = NormalizeConditionalIconSetName(definition.IconSet);
                bool knownIconSet = definition.Source == ExcelConditionalFormattingSource.Office2010Extension
                    ? IsKnownOffice2010ConditionalIconSet(iconSet)
                    : IsKnownStandardConditionalIconSet(iconSet);
                int expectedCount = iconSet?.StartsWith("Three", StringComparison.Ordinal) == true ? 3
                    : iconSet?.StartsWith("Four", StringComparison.Ordinal) == true ? 4
                    : iconSet?.StartsWith("Five", StringComparison.Ordinal) == true ? 5
                    : 0;
                if (!knownIconSet || count != expectedCount) {
                    throw new ArgumentException("Icon-set rules require an icon-set family and three to five thresholds.", nameof(definition));
                }
                if (definition.Source == ExcelConditionalFormattingSource.Standard &&
                    (definition.IconSetCustom == true || customIcons.Count > 0)) {
                    throw new ArgumentException("Custom icon-set rules require the Office 2010 extension surface.", nameof(definition));
                }
                ValidateConditionalFormattingThresholds(iconSetThresholds.Select(item => (item.Type, item.Value)), definition.Source);
                if (definition.IconSetCustom == true && customIcons.Count != count) {
                    throw new ArgumentException("Custom icon-set rules require one custom icon for each threshold.", nameof(definition));
                }
                foreach (ExcelConditionalFormatIcon icon in customIcons) {
                    int iconCount = GetConditionalIconSetCount(NormalizeConditionalIconSetName(icon.IconSet));
                    if (!IsKnownOffice2010ConditionalIconSet(NormalizeConditionalIconSetName(icon.IconSet)) || icon.IconId >= (uint)iconCount) {
                        throw new ArgumentException("Custom icons require a supported icon-set family and an icon id between zero and four.", nameof(definition));
                    }
                }
            }
        }

        private static void ValidateConditionalFormattingThresholds(
            IEnumerable<(string Type, string? Value)> thresholds,
            ExcelConditionalFormattingSource source) {
            foreach ((string type, string? value) in thresholds) {
                string? token = ToConditionalFormatValueToken(type);
                bool known = token is "min" or "max" or "num" or "percent" or "percentile" or "formula"
                    || source == ExcelConditionalFormattingSource.Office2010Extension && (token is "autoMin" or "autoMax");
                if (!known) {
                    throw new ArgumentException($"Conditional-formatting threshold type '{type}' is not supported.");
                }
                if (token is "num" or "percent" or "percentile" or "formula") {
                    if (string.IsNullOrWhiteSpace(value)) {
                        throw new ArgumentException($"Conditional-formatting threshold type '{type}' requires a value.");
                    }
                    XmlConvert.VerifyXmlChars(value!);
                }
            }
        }

        private static string? NormalizeConditionalIconSetName(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return null;
            string normalized = value!.Trim();
            if (normalized[0] == '3') return "Three" + normalized.Substring(1);
            if (normalized[0] == '4') return "Four" + normalized.Substring(1);
            if (normalized[0] == '5') return "Five" + normalized.Substring(1);
            return char.ToUpperInvariant(normalized[0]) + normalized.Substring(1);
        }

        private static bool IsKnownStandardConditionalIconSet(string? value) => value is
            "ThreeArrows" or "ThreeArrowsGray" or "ThreeFlags" or "ThreeTrafficLights1" or
            "ThreeTrafficLights2" or "ThreeSigns" or "ThreeSymbols" or "ThreeSymbols2" or
            "FourArrows" or "FourArrowsGray" or "FourRedToBlack" or "FourRating" or
            "FourTrafficLights" or "FiveArrows" or "FiveArrowsGray" or "FiveRating" or "FiveQuarters";

        private static bool IsKnownOffice2010ConditionalIconSet(string? value) =>
            IsKnownStandardConditionalIconSet(value) || value is "ThreeStars" or "ThreeTriangles" or "FiveBoxes";

        private static int GetConditionalIconSetCount(string? value) =>
            value?.StartsWith("Three", StringComparison.Ordinal) == true ? 3
                : value?.StartsWith("Four", StringComparison.Ordinal) == true ? 4
                : value?.StartsWith("Five", StringComparison.Ordinal) == true ? 5
                : 0;

        private static bool IsKnownConditionalFormattingOperator(string? value) => value is not null && (
            string.Equals(value, "Between", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "NotBetween", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "Equal", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "NotEqual", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "GreaterThan", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "LessThan", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "GreaterThanOrEqual", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "LessThanOrEqual", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "ContainsText", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "NotContains", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "BeginsWith", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "EndsWith", StringComparison.OrdinalIgnoreCase));

        private static bool IsBetweenConditionalFormattingOperator(string? value) =>
            string.Equals(value, "Between", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "NotBetween", StringComparison.OrdinalIgnoreCase);

        private static bool IsKnownConditionalFormattingTimePeriod(string? value) => value is not null && (
            string.Equals(value, "Yesterday", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "Today", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "Tomorrow", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "Last7Days", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "LastWeek", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "ThisWeek", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "NextWeek", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "LastMonth", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "ThisMonth", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "NextMonth", StringComparison.OrdinalIgnoreCase));

        private static bool IsTextConditionalFormattingType(string value) =>
            string.Equals(value, "ContainsText", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "NotContainsText", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "BeginsWith", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "EndsWith", StringComparison.OrdinalIgnoreCase);

        private static void ValidateConditionalFormattingColors(IEnumerable<string> colors) {
            foreach (string color in colors) {
                if (!string.IsNullOrWhiteSpace(color)) NormalizeHexColor(color);
            }
        }

        private static void ValidateConditionalFormattingBorder(ExcelCellBorderSnapshot? border) {
            if (border == null) return;
            ValidateConditionalFormattingBorderSide(border.Left);
            ValidateConditionalFormattingBorderSide(border.Right);
            ValidateConditionalFormattingBorderSide(border.Top);
            ValidateConditionalFormattingBorderSide(border.Bottom);
            ValidateConditionalFormattingBorderSide(border.Diagonal);
        }

        private static void ValidateConditionalFormattingBorderSide(ExcelBorderSideSnapshot? side) {
            if (side == null) return;
            if (NormalizeConditionalFormattingBorderStyle(side.Style) == null) {
                throw new ArgumentException($"Conditional-formatting border style '{side.Style}' is not supported.");
            }
            if (!string.IsNullOrWhiteSpace(side.ColorArgb)) NormalizeHexColor(side.ColorArgb!);
        }

        private static string? NormalizeConditionalFormattingBorderStyle(string? value) =>
            value?.Trim().ToLowerInvariant() switch {
                "none" => "none",
                "thin" => "thin",
                "medium" => "medium",
                "dashed" => "dashed",
                "dotted" => "dotted",
                "thick" => "thick",
                "double" => "double",
                "hair" => "hair",
                "mediumdashed" => "mediumDashed",
                "dashdot" => "dashDot",
                "mediumdashdot" => "mediumDashDot",
                "dashdotdot" => "dashDotDot",
                "mediumdashdotdot" => "mediumDashDotDot",
                "slantdashdot" => "slantDashDot",
                _ => null
            };

        private static void CaptureConditionalFormattingProjectionSignatures(ExcelConditionalFormattingInfo info) {
            info.ProjectedType = info.Type;
            info.ProjectedFormulaSignature = CreateConditionalFormattingFormulaSignature(info);
            info.ProjectedVisualSignature = CreateConditionalFormattingVisualSignature(info);
            info.ProjectedStyleSignature = CreateConditionalFormattingStyleSignature(info);
        }

        private static string CreateConditionalFormattingFormulaSignature(ExcelConditionalFormattingInfo info) =>
            string.Join("\u001f", (info.Formulas ?? Array.Empty<string>()).Select(formula => formula ?? string.Empty));

        private static string CreateConditionalFormattingVisualSignature(ExcelConditionalFormattingInfo info) {
            var values = new List<string> {
                info.Type ?? string.Empty,
                string.Join("\u001e", info.ColorScaleColors ?? Array.Empty<string>()),
                JoinConditionalFormattingThresholds(info.ColorScaleThresholds ?? Array.Empty<ExcelConditionalFormatThreshold>()),
                info.DataBarColor ?? string.Empty,
                JoinConditionalFormattingThresholds(info.DataBarThresholds ?? Array.Empty<ExcelConditionalFormatThreshold>()),
                info.DataBarShowValue.ToString(),
                info.DataBarMinimumLength?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                info.DataBarMaximumLength?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                info.DataBarBorder?.ToString() ?? string.Empty,
                info.DataBarGradient?.ToString() ?? string.Empty,
                info.DataBarDirection ?? string.Empty,
                info.DataBarAxisPosition ?? string.Empty,
                info.DataBarNegativeColorSameAsPositive?.ToString() ?? string.Empty,
                info.DataBarNegativeBorderColorSameAsPositive?.ToString() ?? string.Empty,
                info.DataBarBorderColor ?? string.Empty,
                info.DataBarNegativeColor ?? string.Empty,
                info.DataBarNegativeBorderColor ?? string.Empty,
                info.DataBarAxisColor ?? string.Empty,
                info.IconSet ?? string.Empty,
                info.IconSetShowValue.ToString(),
                info.IconSetReverse.ToString(),
                info.IconSetPercent?.ToString() ?? string.Empty,
                info.IconSetCustom?.ToString() ?? string.Empty,
                string.Join("\u001e", (info.IconSetThresholds ?? Array.Empty<ExcelConditionalIconSetThreshold>()).Select(threshold =>
                    (threshold.Type ?? string.Empty) + "\u001d" + (threshold.Value ?? string.Empty) + "\u001d" + threshold.GreaterThanOrEqual)),
                string.Join("\u001e", (info.CustomIcons ?? Array.Empty<ExcelConditionalFormatIcon>()).Select(icon =>
                    (icon.IconSet ?? string.Empty) + "\u001d" + icon.IconId.ToString(CultureInfo.InvariantCulture)))
            };
            return string.Join("\u001f", values);
        }

        private static string JoinConditionalFormattingThresholds(
            IReadOnlyList<ExcelConditionalFormatThreshold> thresholds) =>
            string.Join("\u001e", thresholds.Select(threshold =>
                (threshold.Type ?? string.Empty) + "\u001d" + (threshold.Value ?? string.Empty)));

        private static string CreateConditionalFormattingStyleSignature(ExcelConditionalFormattingInfo info) =>
            string.Join("\u001f", new[] {
                info.DifferentialFillColorArgb ?? string.Empty,
                info.DifferentialFontColorArgb ?? string.Empty,
                info.DifferentialFontBold?.ToString() ?? string.Empty,
                info.DifferentialFontItalic?.ToString() ?? string.Empty,
                info.DifferentialFontUnderline?.ToString() ?? string.Empty,
                info.DifferentialFontName ?? string.Empty,
                info.DifferentialFontSize?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                CreateConditionalFormattingBorderSignature(info.DifferentialBorder)
            });

        private static string CreateConditionalFormattingBorderSignature(ExcelCellBorderSnapshot? border) {
            if (border == null) return string.Empty;
            return string.Join("\u001e", new[] {
                CreateConditionalFormattingBorderSideSignature(border.Left),
                CreateConditionalFormattingBorderSideSignature(border.Right),
                CreateConditionalFormattingBorderSideSignature(border.Top),
                CreateConditionalFormattingBorderSideSignature(border.Bottom),
                CreateConditionalFormattingBorderSideSignature(border.Diagonal),
                border.DiagonalUp.ToString(),
                border.DiagonalDown.ToString()
            });
        }

        private static string CreateConditionalFormattingBorderSideSignature(ExcelBorderSideSnapshot? side) =>
            side == null ? string.Empty : (side.Style ?? string.Empty) + "\u001d" + (side.ColorArgb ?? string.Empty);

        private static bool IsKnownConditionalFormattingType(string type) =>
            string.Equals(type, "CellIs", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "Expression", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "ColorScale", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "DataBar", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "IconSet", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "Top10", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "UniqueValues", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "DuplicateValues", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "ContainsText", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "NotContainsText", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "BeginsWith", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "EndsWith", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "ContainsBlanks", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "NotContainsBlanks", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "ContainsErrors", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "NotContainsErrors", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "TimePeriod", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(type, "AboveAverage", StringComparison.OrdinalIgnoreCase);
    }
}
