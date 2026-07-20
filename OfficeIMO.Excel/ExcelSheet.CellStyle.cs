using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel {
    public partial class ExcelCell {
        /// <summary>
        /// Gets a read-only snapshot of the cell's visual style.
        /// </summary>
        public ExcelCellStyleSnapshot GetStyle() => Sheet.GetCellStyle(Row, Column);
    }

    public partial class ExcelSheet {
        /// <summary>
        /// Gets a read-only snapshot of the visual style assigned to a worksheet cell.
        /// </summary>
        public ExcelCellStyleSnapshot GetCellStyle(int row, int column) {
            Cell? cell = TryGetExistingCell(row, column);
            if (cell == null) {
                return new ExcelCellStyleSnapshot();
            }

            WorkbookPart? workbookPart = _excelDocument.WorkbookPartRoot;
            Stylesheet? stylesheet = workbookPart?.WorkbookStylesPart?.Stylesheet;
            if (stylesheet == null) {
                return new ExcelCellStyleSnapshot {
                    StyleIndex = cell.StyleIndex?.Value ?? 0U
                };
            }

            CellFormat format = GetBaseCellFormat(stylesheet, cell.StyleIndex?.Value ?? 0U);
            Font? font = stylesheet.Fonts?.Elements<Font>().ElementAtOrDefault((int)(format.FontId?.Value ?? 0U));
            Fill? fill = stylesheet.Fills?.Elements<Fill>().ElementAtOrDefault((int)(format.FillId?.Value ?? 0U));
            Border? border = stylesheet.Borders?.Elements<Border>().ElementAtOrDefault((int)(format.BorderId?.Value ?? 0U));
            uint numberFormatId = format.NumberFormatId?.Value ?? 0U;
            string? numberFormatCode = GetNumberFormatCode(stylesheet, numberFormatId);
            bool hasSimpleGradient = ExcelGradientFillResolver.TryResolveSimpleLinearGradient(fill, workbookPart, out ExcelGradientFillInfo gradient);

            return new ExcelCellStyleSnapshot {
                StyleIndex = cell.StyleIndex?.Value ?? 0U,
                NumberFormatId = numberFormatId,
                NumberFormatCode = numberFormatCode,
                IsDateLike = IsBuiltInDate(numberFormatId) || ExcelNumberFormatClassifier.LooksLikeDateFormat(numberFormatCode),
                Bold = font?.Bold != null,
                Italic = font?.Italic != null,
                Underline = font?.Underline != null,
                FontName = font?.FontName?.Val?.Value,
                IsFontFamilyExplicit = (format.FontId?.Value ?? 0U) != 0U,
                FontSize = font?.FontSize?.Val?.Value,
                FontColorArgb = ExcelThemeColorResolver.Resolve(font?.Color, workbookPart),
                FillColorArgb = GetFillArgb(fill, workbookPart),
                FillPatternType = GetFillPatternType(fill),
                FillPatternForegroundColorArgb = GetFillPatternForegroundArgb(fill, workbookPart),
                FillPatternBackgroundColorArgb = GetFillPatternBackgroundArgb(fill, workbookPart),
                FillGradientUnsupported = fill?.GradientFill != null && !hasSimpleGradient,
                FillGradientStartColorArgb = hasSimpleGradient ? gradient.StartColorArgb : null,
                FillGradientEndColorArgb = hasSimpleGradient ? gradient.EndColorArgb : null,
                FillGradientStops = hasSimpleGradient ? CreateGradientStopSnapshots(gradient) : Array.Empty<ExcelGradientFillStopSnapshot>(),
                FillGradientDegree = hasSimpleGradient ? gradient.Degree : null,
                Border = BuildBorderSnapshot(border, workbookPart),
                HorizontalAlignment = format.Alignment?.Horizontal?.InnerText,
                VerticalAlignment = format.Alignment?.Vertical?.InnerText,
                TextRotation = ToTextRotation(format.Alignment?.TextRotation?.Value),
                TextIndent = format.Alignment?.Indent?.Value,
                WrapText = format.Alignment?.WrapText?.Value == true,
                ShrinkToFit = format.Alignment?.ShrinkToFit?.Value == true
            };
        }

        private static IReadOnlyList<ExcelGradientFillStopSnapshot> CreateGradientStopSnapshots(ExcelGradientFillInfo gradient) =>
            gradient.Stops.Select(stop => new ExcelGradientFillStopSnapshot(stop.Offset, stop.ColorArgb)).ToArray();

        private static int? ToTextRotation(uint? value) {
            if (!value.HasValue) {
                return null;
            }

            return value.Value <= int.MaxValue ? (int)value.Value : (int?)null;
        }

        private static string? GetFillArgb(Fill? fill, WorkbookPart? workbookPart) {
            PatternFill? pattern = fill?.PatternFill;
            if (pattern == null) {
                return null;
            }

            if (pattern.PatternType?.Value == PatternValues.Solid) {
                return ExcelThemeColorResolver.Resolve(pattern.ForegroundColor, workbookPart)
                    ?? ExcelThemeColorResolver.Resolve(pattern.BackgroundColor, workbookPart);
            }

            return ExcelThemeColorResolver.Resolve(pattern.BackgroundColor, workbookPart);
        }

        private static string? GetFillPatternType(Fill? fill) {
            PatternFill? pattern = fill?.PatternFill;
            if (pattern?.PatternType?.Value == null) {
                return fill?.GradientFill != null ? "gradient" : null;
            }

            if (pattern.PatternType.Value == PatternValues.None) {
                return null;
            }

            return NormalizePatternType(pattern.PatternType.InnerText);
        }

        private static string? GetFillPatternForegroundArgb(Fill? fill, WorkbookPart? workbookPart) {
            PatternFill? pattern = fill?.PatternFill;
            return pattern == null ? null : ExcelThemeColorResolver.Resolve(pattern.ForegroundColor, workbookPart);
        }

        private static string? GetFillPatternBackgroundArgb(Fill? fill, WorkbookPart? workbookPart) {
            PatternFill? pattern = fill?.PatternFill;
            return pattern == null ? null : ExcelThemeColorResolver.Resolve(pattern.BackgroundColor, workbookPart);
        }

        private static string NormalizePatternType(string? value) {
            string text = value ?? string.Empty;
            return string.IsNullOrEmpty(text) ? string.Empty : char.ToLowerInvariant(text[0]) + text.Substring(1);
        }

        private static ExcelCellBorderSnapshot? BuildBorderSnapshot(Border? border, WorkbookPart? workbookPart) {
            if (border == null) {
                return null;
            }

            ExcelBorderSideSnapshot? left = BuildBorderSideSnapshot(border.LeftBorder, workbookPart);
            ExcelBorderSideSnapshot? right = BuildBorderSideSnapshot(border.RightBorder, workbookPart);
            ExcelBorderSideSnapshot? top = BuildBorderSideSnapshot(border.TopBorder, workbookPart);
            ExcelBorderSideSnapshot? bottom = BuildBorderSideSnapshot(border.BottomBorder, workbookPart);
            ExcelBorderSideSnapshot? diagonal = BuildBorderSideSnapshot(border.DiagonalBorder, workbookPart);
            bool diagonalUp = border.DiagonalUp?.Value == true;
            bool diagonalDown = border.DiagonalDown?.Value == true;
            if (left == null && right == null && top == null && bottom == null && (!diagonalUp && !diagonalDown || diagonal == null)) {
                return null;
            }

            return new ExcelCellBorderSnapshot {
                Left = left,
                Right = right,
                Top = top,
                Bottom = bottom,
                Diagonal = diagonal,
                DiagonalUp = diagonalUp,
                DiagonalDown = diagonalDown
            };
        }

        private static ExcelBorderSideSnapshot? BuildBorderSideSnapshot(BorderPropertiesType? borderSide, WorkbookPart? workbookPart) {
            if (borderSide == null) {
                return null;
            }

            string? style = ExtractBorderStyle(borderSide);
            string? colorArgb = ExcelThemeColorResolver.Resolve(borderSide.GetFirstChild<Color>(), workbookPart);
            if (string.IsNullOrWhiteSpace(style) && string.IsNullOrWhiteSpace(colorArgb)) {
                return null;
            }

            return new ExcelBorderSideSnapshot {
                Style = style,
                ColorArgb = colorArgb
            };
        }

        private static string? ExtractBorderStyle(BorderPropertiesType borderSide) {
            string xml = borderSide.OuterXml;
            if (string.IsNullOrWhiteSpace(xml)) {
                return null;
            }

            const string marker = "style=\"";
            int index = xml.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (index < 0) {
                return null;
            }

            index += marker.Length;
            int endIndex = xml.IndexOf('"', index);
            if (endIndex <= index) {
                return null;
            }

            string value = xml.Substring(index, endIndex - index);
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim().ToLowerInvariant();
        }

        private static string? GetNumberFormatCode(Stylesheet stylesheet, uint numberFormatId) {
            if (stylesheet.NumberingFormats != null) {
                foreach (NumberingFormat numberingFormat in stylesheet.NumberingFormats.Elements<NumberingFormat>()) {
                    if (numberingFormat.NumberFormatId?.Value == numberFormatId) {
                        return numberingFormat.FormatCode?.Value;
                    }
                }
            }

            return ExcelBuiltInNumberFormats.GetCode(numberFormatId);
        }

        private static bool IsBuiltInDate(uint numberFormatId) =>
            ExcelBuiltInNumberFormats.IsDate(numberFormatId);

    }
}
