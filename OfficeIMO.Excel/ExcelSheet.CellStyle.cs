using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

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

            return new ExcelCellStyleSnapshot {
                StyleIndex = cell.StyleIndex?.Value ?? 0U,
                NumberFormatId = numberFormatId,
                NumberFormatCode = numberFormatCode,
                IsDateLike = IsBuiltInDate(numberFormatId) || ExcelNumberFormatClassifier.LooksLikeDateFormat(numberFormatCode),
                Bold = font?.Bold != null,
                Italic = font?.Italic != null,
                Underline = font?.Underline != null,
                FontColorArgb = NormalizeReadArgb(font?.Color?.Rgb?.Value),
                FillColorArgb = GetFillArgb(fill),
                Border = BuildBorderSnapshot(border),
                HorizontalAlignment = format.Alignment?.Horizontal?.InnerText,
                VerticalAlignment = format.Alignment?.Vertical?.InnerText,
                WrapText = format.Alignment?.WrapText?.Value == true
            };
        }

        private static string? GetFillArgb(Fill? fill) {
            PatternFill? pattern = fill?.PatternFill;
            if (pattern == null || pattern.PatternType?.Value != PatternValues.Solid) {
                return null;
            }

            return NormalizeReadArgb(pattern.ForegroundColor?.Rgb?.Value)
                ?? NormalizeReadArgb(pattern.BackgroundColor?.Rgb?.Value);
        }

        private static ExcelCellBorderSnapshot? BuildBorderSnapshot(Border? border) {
            if (border == null) {
                return null;
            }

            ExcelBorderSideSnapshot? left = BuildBorderSideSnapshot(border.LeftBorder);
            ExcelBorderSideSnapshot? right = BuildBorderSideSnapshot(border.RightBorder);
            ExcelBorderSideSnapshot? top = BuildBorderSideSnapshot(border.TopBorder);
            ExcelBorderSideSnapshot? bottom = BuildBorderSideSnapshot(border.BottomBorder);
            ExcelBorderSideSnapshot? diagonal = BuildBorderSideSnapshot(border.DiagonalBorder);
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

        private static ExcelBorderSideSnapshot? BuildBorderSideSnapshot(BorderPropertiesType? borderSide) {
            if (borderSide == null) {
                return null;
            }

            string? style = ExtractBorderStyle(borderSide);
            string? colorArgb = NormalizeReadArgb(borderSide.GetFirstChild<Color>()?.Rgb?.Value);
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

            return GetBuiltInNumberFormatCode(numberFormatId);
        }

        private static bool IsBuiltInDate(uint numberFormatId) =>
            numberFormatId is 14 or 15 or 16 or 17 or 18 or 19 or 20 or 21 or 22
                or 27 or 30 or 36 or 45 or 46 or 47;

        private static string? NormalizeReadArgb(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string hex = value!.Trim();
            if (hex.StartsWith("#", StringComparison.Ordinal)) {
                hex = hex.Substring(1);
            }

            if (hex.Length == 6) {
                hex = "FF" + hex;
            } else if (hex.Length != 8) {
                return null;
            }

            for (int i = 0; i < hex.Length; i++) {
                char ch = hex[i];
                bool isHex = (ch >= '0' && ch <= '9') ||
                    (ch >= 'a' && ch <= 'f') ||
                    (ch >= 'A' && ch <= 'F');
                if (!isHex) {
                    return null;
                }
            }

            return hex.ToUpperInvariant();
        }
    }
}
