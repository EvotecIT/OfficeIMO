using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Utilities;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const string Office2010ConditionalFormattingExtensionUri = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}";

        private bool HasOffice2010ConditionalFormatting() =>
            WorksheetRoot.GetFirstChild<WorksheetExtensionList>()?
                .Descendants<X14.ConditionalFormatting>()
                .Any() == true;

        private bool ClearOffice2010ConditionalFormattingCore(string? a1Range) {
            WorksheetExtensionList? extensions = WorksheetRoot.GetFirstChild<WorksheetExtensionList>();
            if (extensions == null) return false;
            List<X14.ConditionalFormatting> existing = extensions
                .Descendants<X14.ConditionalFormatting>()
                .ToList();
            if (existing.Count == 0) return false;

            if (string.IsNullOrWhiteSpace(a1Range)) {
                foreach (X14.ConditionalFormatting formatting in existing) formatting.Remove();
            } else {
                RemoveOffice2010ConditionalFormattingOverlap(ParseReferenceArgument(a1Range!));
            }

            foreach (X14.ConditionalFormattings formattings in
                WorksheetRoot.Descendants<X14.ConditionalFormattings>().ToList()) {
                if (!formattings.Elements<X14.ConditionalFormatting>().Any()) formattings.Remove();
            }
            CleanupEmptyMetadataExtensions();
            return existing.Count != (WorksheetRoot.GetFirstChild<WorksheetExtensionList>()?
                .Descendants<X14.ConditionalFormatting>().Count() ?? 0);
        }

        private void CollectOffice2010ConditionalFormattingRules(
            (int r1, int c1, int r2, int c2)? filter,
            int maximumRules,
            int maximumDiscoveryItems,
            List<ExcelConditionalFormattingInfo> list,
            SortedSet<ConditionalFormattingCandidate>? retained,
            ref long ruleOrder,
            ref int containersExamined,
            ref bool truncated) {
            WorksheetExtensionList? extensions = WorksheetRoot.GetFirstChild<WorksheetExtensionList>();
            if (extensions == null) return;
            foreach (X14.ConditionalFormatting conditional in
                extensions.Descendants<X14.ConditionalFormatting>()) {
                if (retained != null && containersExamined++ >= maximumDiscoveryItems) {
                    truncated = true;
                    return;
                }

                string range = conditional.GetFirstChild<Xm.ReferenceSequence>()?.Text ?? string.Empty;
                if (filter.HasValue && !string.IsNullOrWhiteSpace(range)) {
                    long maximumRangeCharacters = (long)maximumDiscoveryItems * 64L;
                    if (range.Length > maximumRangeCharacters) {
                        truncated = true;
                        return;
                    }

                    if (!ReferenceListOverlaps(
                        range,
                        filter.Value,
                        maximumDiscoveryItems,
                        out bool referencesTruncated)) {
                        if (referencesTruncated) {
                            truncated = true;
                            return;
                        }
                        continue;
                    }
                }

                foreach (X14.ConditionalFormattingRule rule in
                    conditional.Elements<X14.ConditionalFormattingRule>()) {
                    if (retained == null) {
                        list.Add(ReadOffice2010ConditionalFormattingInfo(
                            rule,
                            range,
                            _excelDocument.WorkbookPartRoot));
                        continue;
                    }

                    var candidate = new ConditionalFormattingCandidate(
                        rule,
                        range,
                        NormalizeConditionalFormattingPriority(rule),
                        ruleOrder++);
                    if (retained.Count < maximumRules) {
                        retained.Add(candidate);
                        continue;
                    }

                    truncated = true;
                    ConditionalFormattingCandidate worst = retained.Max;
                    if (ConditionalFormattingCandidateComparer.Instance.Compare(candidate, worst) < 0) {
                        retained.Remove(worst);
                        retained.Add(candidate);
                    }
                    return;
                }
            }
        }

        private static int ReadConditionalFormattingPriority(OpenXmlElement rule) {
            if (rule is ConditionalFormattingRule standardRule) {
                return (int)(standardRule.Priority?.Value ?? 0);
            }

            if (rule is X14.ConditionalFormattingRule extensionRule) {
                return (int)(extensionRule.Priority?.Value ?? 0);
            }

            return 0;
        }

        private ExcelConditionalFormattingInfo ReadOffice2010ConditionalFormattingInfo(
            OpenXmlElement element,
            string range,
            WorkbookPart workbookPart) {
            if (element is not X14.ConditionalFormattingRule rule) {
                throw new InvalidOperationException("The conditional-formatting rule is not an Office 2010 extension rule.");
            }

            X14.DifferentialType? differential = rule.GetFirstChild<X14.DifferentialType>();
            X14.ColorScale? colorScale = rule.GetFirstChild<X14.ColorScale>();
            X14.DataBar? dataBar = rule.GetFirstChild<X14.DataBar>();
            X14.IconSet? iconSet = rule.GetFirstChild<X14.IconSet>();
            PatternFill? pattern = differential?.Fill?.PatternFill;
            Font? font = differential?.Font;

            var info = new ExcelConditionalFormattingInfo {
                Source = ExcelConditionalFormattingSource.Office2010Extension,
                ExtensionId = rule.Id?.Value,
                OwnerSheet = this,
                OwnerContainer = rule.Parent,
                BackingRule = rule,
                HasPreservedUnknownMarkup = HasUnprojectedConditionalFormattingMarkup(rule, rule.Parent),
                Range = range,
                Type = NormalizeConditionalFormatType(rule.Type?.InnerText),
                Operator = NormalizeConditionalFormatOperator(rule.Operator?.InnerText),
                Text = rule.Text?.Value,
                TimePeriod = NormalizeConditionalTimePeriod(rule.TimePeriod?.InnerText),
                Priority = (int)(rule.Priority?.Value ?? 0),
                StopIfTrue = rule.StopIfTrue?.Value ?? false,
                DifferentialFillColorArgb = ExcelThemeColorResolver.Resolve(pattern?.ForegroundColor, workbookPart)
                    ?? ExcelThemeColorResolver.Resolve(pattern?.BackgroundColor, workbookPart),
                DifferentialFontColorArgb = ExcelThemeColorResolver.Resolve(font?.Color, workbookPart),
                DifferentialFontBold = ReadDifferentialBoolean(font?.Bold),
                DifferentialFontItalic = ReadDifferentialBoolean(font?.Italic),
                DifferentialFontUnderline = ReadDifferentialBoolean(font?.Underline),
                DifferentialFontName = font?.FontName?.Val?.Value,
                DifferentialFontSize = font?.FontSize?.Val?.Value,
                DifferentialBorder = RemoveDifferentialColorOnlyBorderSides(
                    BuildBorderSnapshot(differential?.Border, workbookPart)),
                Formulas = rule.Elements<Xm.Formula>().Select(formula => formula.Text ?? string.Empty).ToArray(),
                ColorScaleColors = colorScale?.Elements<X14.Color>()
                    .Select(color => ReadOffice2010Color(color) ?? string.Empty)
                    .ToArray() ?? Array.Empty<string>(),
                ColorScaleThresholds = ReadOffice2010Thresholds(colorScale),
                DataBarColor = ReadOffice2010Color(dataBar?.GetFirstChild<X14.FillColor>()),
                DataBarThresholds = ReadOffice2010Thresholds(dataBar),
                DataBarShowValue = dataBar?.ShowValue?.Value ?? true,
                DataBarMinimumLength = dataBar?.MinLength?.Value,
                DataBarMaximumLength = dataBar?.MaxLength?.Value,
                DataBarBorder = dataBar?.Border?.Value,
                DataBarGradient = dataBar?.Gradient?.Value,
                DataBarDirection = dataBar?.Direction?.InnerText,
                DataBarAxisPosition = dataBar?.AxisPosition?.InnerText,
                DataBarNegativeColorSameAsPositive = dataBar?.NegativeBarColorSameAsPositive?.Value,
                DataBarNegativeBorderColorSameAsPositive = dataBar?.NegativeBarBorderColorSameAsPositive?.Value,
                DataBarBorderColor = ReadOffice2010Color(dataBar?.GetFirstChild<X14.BorderColor>()),
                DataBarNegativeColor = ReadOffice2010Color(dataBar?.GetFirstChild<X14.NegativeFillColor>()),
                DataBarNegativeBorderColor = ReadOffice2010Color(dataBar?.GetFirstChild<X14.NegativeBorderColor>()),
                DataBarAxisColor = ReadOffice2010Color(dataBar?.GetFirstChild<X14.BarAxisColor>()),
                IconSet = NormalizeConditionalIconSetName(iconSet?.IconSetTypes?.InnerText),
                IconSetShowValue = iconSet?.ShowValue?.Value ?? true,
                IconSetReverse = iconSet?.Reverse?.Value ?? false,
                IconSetPercent = iconSet?.Percent?.Value,
                IconSetCustom = iconSet?.Custom?.Value,
                IconSetThresholds = ReadOffice2010IconThresholds(iconSet),
                CustomIcons = iconSet?.Elements<X14.ConditionalFormattingIcon>()
                    .Select(icon => new ExcelConditionalFormatIcon {
                        IconSet = NormalizeConditionalIconSetName(icon.IconSet?.InnerText) ?? string.Empty,
                        IconId = icon.IconId?.Value ?? 0U
                    })
                    .ToArray() ?? Array.Empty<ExcelConditionalFormatIcon>(),
                TopBottomRank = rule.Rank?.Value,
                TopBottomBottom = rule.Bottom?.Value ?? false,
                TopBottomPercent = rule.Percent?.Value ?? false,
                AboveAverageAbove = rule.AboveAverage?.Value ?? true,
                AboveAverageEqual = rule.EqualAverage?.Value ?? false,
                AboveAverageStdDev = rule.StandardDeviation?.Value
            };
            CaptureConditionalFormattingProjectionSignatures(info);
            return info;
        }

        private static string NormalizeConditionalFormatType(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            if (string.Equals(value, "cellIs", StringComparison.OrdinalIgnoreCase)) return "CellIs";
            if (string.Equals(value, "expression", StringComparison.OrdinalIgnoreCase)) return "Expression";
            if (string.Equals(value, "colorScale", StringComparison.OrdinalIgnoreCase)) return "ColorScale";
            if (string.Equals(value, "dataBar", StringComparison.OrdinalIgnoreCase)) return "DataBar";
            if (string.Equals(value, "iconSet", StringComparison.OrdinalIgnoreCase)) return "IconSet";
            if (string.Equals(value, "top10", StringComparison.OrdinalIgnoreCase)) return "Top10";
            if (string.Equals(value, "uniqueValues", StringComparison.OrdinalIgnoreCase)) return "UniqueValues";
            if (string.Equals(value, "duplicateValues", StringComparison.OrdinalIgnoreCase)) return "DuplicateValues";
            if (string.Equals(value, "containsText", StringComparison.OrdinalIgnoreCase)) return "ContainsText";
            if (string.Equals(value, "notContainsText", StringComparison.OrdinalIgnoreCase)) return "NotContainsText";
            if (string.Equals(value, "beginsWith", StringComparison.OrdinalIgnoreCase)) return "BeginsWith";
            if (string.Equals(value, "endsWith", StringComparison.OrdinalIgnoreCase)) return "EndsWith";
            if (string.Equals(value, "containsBlanks", StringComparison.OrdinalIgnoreCase)) return "ContainsBlanks";
            if (string.Equals(value, "notContainsBlanks", StringComparison.OrdinalIgnoreCase)) return "NotContainsBlanks";
            if (string.Equals(value, "containsErrors", StringComparison.OrdinalIgnoreCase)) return "ContainsErrors";
            if (string.Equals(value, "notContainsErrors", StringComparison.OrdinalIgnoreCase)) return "NotContainsErrors";
            if (string.Equals(value, "timePeriod", StringComparison.OrdinalIgnoreCase)) return "TimePeriod";
            if (string.Equals(value, "aboveAverage", StringComparison.OrdinalIgnoreCase)) return "AboveAverage";
            return value!;
        }

        private static string? NormalizeConditionalFormatOperator(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return null;
            return char.ToUpperInvariant(value![0]) + value.Substring(1);
        }

        private static string? NormalizeConditionalTimePeriod(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return null;
            return char.ToUpperInvariant(value![0]) + value.Substring(1);
        }

        private static IReadOnlyList<ExcelConditionalFormatThreshold> ReadOffice2010Thresholds(
            OpenXmlElement? parent) {
            if (parent == null) return Array.Empty<ExcelConditionalFormatThreshold>();
            return parent.Elements<X14.ConditionalFormattingValueObject>()
                .Select(threshold => new ExcelConditionalFormatThreshold {
                    Type = threshold.Type?.InnerText ?? string.Empty,
                    Value = threshold.Formula?.Text
                })
                .ToArray();
        }

        private static IReadOnlyList<ExcelConditionalIconSetThreshold> ReadOffice2010IconThresholds(
            X14.IconSet? iconSet) {
            if (iconSet == null) return Array.Empty<ExcelConditionalIconSetThreshold>();
            return iconSet.Elements<X14.ConditionalFormattingValueObject>()
                .Select(threshold => new ExcelConditionalIconSetThreshold {
                    Type = threshold.Type?.InnerText ?? string.Empty,
                    Value = threshold.Formula?.Text,
                    GreaterThanOrEqual = threshold.GreaterThanOrEqual?.Value ?? true
                })
                .ToArray();
        }

        private static string? ReadOffice2010Color(X14.ColorType? color) => color?.Rgb?.Value;

        private static bool HasUnprojectedConditionalFormattingMarkup(
            OpenXmlElement rule,
            OpenXmlElement? container) {
            if (rule.ExtendedAttributes.Any() || rule.Descendants<OpenXmlUnknownElement>().Any()) return true;
            if (container != null &&
                (container.ExtendedAttributes.Any() ||
                 container.ChildElements.Any(child => child is OpenXmlUnknownElement ||
                     child is DocumentFormat.OpenXml.Spreadsheet.ExtensionList || child is X14.ExtensionList))) return true;
            if (rule is ConditionalFormattingRule standard) {
                return standard.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.ExtensionList>() != null;
            }
            if (rule is X14.ConditionalFormattingRule extension) {
                return extension.GetFirstChild<X14.ExtensionList>() != null;
            }
            return true;
        }
    }
}
