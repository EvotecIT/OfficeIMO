using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static CellFormat GetBaseCellFormat(Stylesheet stylesheet, uint styleIndex) {
            var cellFormats = stylesheet.CellFormats?.Elements<CellFormat>().ToList();
            var baseFormat = cellFormats?.ElementAtOrDefault((int)styleIndex);
            if (baseFormat != null) {
                return (CellFormat)baseFormat.CloneNode(true);
            }

            return new CellFormat {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U
            };
        }

        private static void ApplyCellFormatOverride(Stylesheet stylesheet, Cell cell, Action<CellFormat> mutate) {
            var baseFormat = GetBaseCellFormat(stylesheet, cell.StyleIndex?.Value ?? 0U);
            mutate(baseFormat);
            cell.StyleIndex = AppendOrReuseCellFormat(stylesheet, baseFormat);
        }

        private static uint GetOrAddCellFormatOverride(
            Dictionary<uint, uint> styleIndexes,
            Stylesheet stylesheet,
            uint baseStyleIndex,
            Action<CellFormat> mutate) {
            if (!styleIndexes.TryGetValue(baseStyleIndex, out uint styleIndex)) {
                var format = GetBaseCellFormat(stylesheet, baseStyleIndex);
                mutate(format);
                styleIndex = AppendOrReuseCellFormat(stylesheet, format);
                styleIndexes.Add(baseStyleIndex, styleIndex);
            }

            return styleIndex;
        }

        private static uint AppendOrReuseCellFormat(Stylesheet stylesheet, CellFormat candidate) {
            var cellFormats = stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            var existing = cellFormats.Elements<CellFormat>()
                .Select((format, index) => new { format, index })
                .FirstOrDefault(entry => string.Equals(entry.format.OuterXml, candidate.OuterXml, StringComparison.Ordinal));
            if (existing != null) {
                return (uint)existing.index;
            }

            cellFormats.Append(candidate);
            cellFormats.Count = (uint)cellFormats.Count();
            return cellFormats.Count!.Value - 1;
        }

        private static uint GetOrCreateFill(Stylesheet stylesheet, Fill candidate) {
            var fills = stylesheet.Fills ??= new Fills();
            var existing = fills.Elements<Fill>()
                .Select((fill, index) => new { fill, index })
                .FirstOrDefault(entry => string.Equals(entry.fill.OuterXml, candidate.OuterXml, StringComparison.Ordinal));
            if (existing != null) {
                return (uint)existing.index;
            }

            fills.Append(candidate);
            fills.Count = (uint)fills.Count();
            return fills.Count!.Value - 1;
        }

        private static uint GetOrCreateNumberFormatId(Stylesheet stylesheet, string numberFormat) {
            stylesheet.NumberingFormats ??= new NumberingFormats();
            NumberingFormat? existingFormat = stylesheet.NumberingFormats.Elements<NumberingFormat>()
                .FirstOrDefault(n => n.FormatCode != null && n.FormatCode.Value == numberFormat);

            if (existingFormat != null) {
                return existingFormat.NumberFormatId!.Value;
            }

            uint numberFormatId = stylesheet.NumberingFormats.Elements<NumberingFormat>().Any()
                ? stylesheet.NumberingFormats.Elements<NumberingFormat>().Max(n => n.NumberFormatId!.Value) + 1
                : 164U;
            NumberingFormat numberingFormat = new NumberingFormat {
                NumberFormatId = numberFormatId,
                FormatCode = StringValue.FromString(numberFormat)
            };
            stylesheet.NumberingFormats.Append(numberingFormat);
            stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Count();
            return numberFormatId;
        }

        private static uint GetOrCreateBorderVariant(Stylesheet stylesheet, uint? baseBorderId, Action<Border> mutate) {
            var borders = stylesheet.Borders ??= new Borders(new Border());
            var baseBorder = borders.Elements<Border>().ElementAtOrDefault((int)(baseBorderId ?? 0U));
            var candidate = baseBorder != null
                ? (Border)baseBorder.CloneNode(true)
                : new Border();

            mutate(candidate);

            var existing = borders.Elements<Border>()
                .Select((border, index) => new { border, index })
                .FirstOrDefault(entry => string.Equals(entry.border.OuterXml, candidate.OuterXml, StringComparison.Ordinal));
            if (existing != null) {
                return (uint)existing.index;
            }

            borders.Append(candidate);
            borders.Count = (uint)borders.Count();
            return borders.Count!.Value - 1;
        }

        private static uint GetOrCreateFontVariant(Stylesheet stylesheet, uint? baseFontId, Action<DocumentFormat.OpenXml.Spreadsheet.Font> mutate) {
            var fonts = stylesheet.Fonts ??= new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font());
            var baseFont = fonts.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAtOrDefault((int)(baseFontId ?? 0U));
            var candidate = baseFont != null
                ? (DocumentFormat.OpenXml.Spreadsheet.Font)baseFont.CloneNode(true)
                : new DocumentFormat.OpenXml.Spreadsheet.Font();

            mutate(candidate);

            var existing = fonts.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>()
                .Select((font, index) => new { font, index })
                .FirstOrDefault(entry => string.Equals(entry.font.OuterXml, candidate.OuterXml, StringComparison.Ordinal));
            if (existing != null) {
                return (uint)existing.index;
            }

            fonts.Append(candidate);
            fonts.Count = (uint)fonts.Count();
            return fonts.Count!.Value - 1;
        }

        private static uint? GetOptionalValue(UInt32Value? value) {
            return value != null ? value.Value : (uint?)null;
        }

        private static void SetBold(DocumentFormat.OpenXml.Spreadsheet.Font font, bool bold) {
            font.Bold = bold ? new Bold() : null;
        }

        private static void SetItalic(DocumentFormat.OpenXml.Spreadsheet.Font font, bool italic) {
            font.Italic = italic ? new Italic() : null;
        }

        private static void SetUnderline(DocumentFormat.OpenXml.Spreadsheet.Font font, bool underline) {
            font.Underline = underline ? new Underline() : null;
        }

        private static void SetFontColor(DocumentFormat.OpenXml.Spreadsheet.Font font, string argb) {
            font.Color = new DocumentFormat.OpenXml.Spreadsheet.Color {
                Rgb = argb
            };
        }

        private static void SetFontName(DocumentFormat.OpenXml.Spreadsheet.Font font, string fontName) {
            font.FontName = new FontName {
                Val = fontName.Trim()
            };
        }

        private static void SetFontSize(DocumentFormat.OpenXml.Spreadsheet.Font font, double fontSize) {
            font.FontSize = new FontSize {
                Val = fontSize
            };
        }

        private static void SetUniformBorder(Border border, BorderStyleValues style, string? hexColor) {
            var argb = string.IsNullOrWhiteSpace(hexColor) ? null : NormalizeHexColor(hexColor!);
            border.LeftBorder = CreateBorderSide<LeftBorder>(style, argb);
            border.RightBorder = CreateBorderSide<RightBorder>(style, argb);
            border.TopBorder = CreateBorderSide<TopBorder>(style, argb);
            border.BottomBorder = CreateBorderSide<BottomBorder>(style, argb);
        }

        private static void SetDiagonalBorder(Border border, BorderStyleValues style, string? hexColor, bool diagonalUp, bool diagonalDown) {
            var argb = string.IsNullOrWhiteSpace(hexColor) ? null : NormalizeHexColor(hexColor!);
            border.DiagonalBorder = CreateBorderSide<DiagonalBorder>(style, argb);
            border.DiagonalUp = diagonalUp;
            border.DiagonalDown = diagonalDown;
        }

        private static T CreateBorderSide<T>(BorderStyleValues style, string? argb) where T : BorderPropertiesType, new() {
            var side = new T {
                Style = style
            };

            if (!string.IsNullOrWhiteSpace(argb)) {
                side.Append(new Color {
                    Rgb = argb
                });
            }

            return side;
        }

        /// <summary>
        /// Ensures required default style primitives exist and their counts are consistent.
        /// Excel expects at least 1 Font, 2 Fills (None, Gray125), 1 Border,
        /// 1 CellStyleFormat, and 1 CellFormat present.
        /// </summary>
        private static void EnsureDefaultStylePrimitives(Stylesheet stylesheet) {
            // Fonts
            if (stylesheet.Fonts == null || !stylesheet.Fonts.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().Any()) {
                stylesheet.Fonts = new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font(new FontSize { Val = 11D }, new FontName { Val = "Calibri" }));
            } else {
                var defaultFont = stylesheet.Fonts.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().First();
                defaultFont.FontSize ??= new FontSize { Val = 11D };
                defaultFont.FontName ??= new FontName { Val = "Calibri" };
            }
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

            // Fills: ensure index 0 = None, index 1 = Gray125
            if (stylesheet.Fills == null) {
                stylesheet.Fills = new Fills();
            }
            var fills = stylesheet.Fills.Elements<Fill>().ToList();
            bool hasNone = fills.Any(f => f.PatternFill?.PatternType?.Value == PatternValues.None);
            bool hasGray = fills.Any(f => f.PatternFill?.PatternType?.Value == PatternValues.Gray125);
            if (!hasNone) {
                stylesheet.Fills.AppendChild(new Fill(new PatternFill { PatternType = PatternValues.None }));
            }
            if (!hasGray) {
                stylesheet.Fills.AppendChild(new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
            }
            stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();

            // Borders
            if (stylesheet.Borders == null || !stylesheet.Borders.Elements<Border>().Any()) {
                stylesheet.Borders = new Borders(new Border());
            }
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();

            // Cell style formats
            if (stylesheet.CellStyleFormats == null || !stylesheet.CellStyleFormats.Elements<CellFormat>().Any()) {
                stylesheet.CellStyleFormats = new CellStyleFormats(new CellFormat {
                    NumberFormatId = 0U,
                    FontId = 0U,
                    FillId = 0U,
                    BorderId = 0U
                });
            }
            stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.Count();

            // Cell formats
            if (stylesheet.CellFormats == null || !stylesheet.CellFormats.Elements<CellFormat>().Any()) {
                stylesheet.CellFormats = new CellFormats(new CellFormat {
                    NumberFormatId = 0U,
                    FontId = 0U,
                    FillId = 0U,
                    BorderId = 0U,
                    FormatId = 0U
                });
            }
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();

            if (stylesheet.CellStyles == null || !stylesheet.CellStyles.Elements<CellStyle>().Any()) {
                stylesheet.CellStyles = new CellStyles(new CellStyle {
                    Name = "Normal",
                    FormatId = 0U,
                    BuiltinId = 0U
                });
            }
            stylesheet.CellStyles.Count = (uint)stylesheet.CellStyles.Count();

            stylesheet.DifferentialFormats ??= new DifferentialFormats();
            stylesheet.DifferentialFormats.Count = (uint)stylesheet.DifferentialFormats.Count();

            stylesheet.TableStyles ??= new TableStyles {
                DefaultTableStyle = "TableStyleMedium2",
                DefaultPivotStyle = "PivotStyleLight16"
            };
            stylesheet.TableStyles.Count = (uint)stylesheet.TableStyles.Count();

            // Numbering formats count normalization
            if (stylesheet.NumberingFormats != null) {
                stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Count();
            }
        }
    }
}
