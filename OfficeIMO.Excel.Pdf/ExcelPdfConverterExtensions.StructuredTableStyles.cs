using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static IReadOnlyList<StructuredTableVisualData> ReadStructuredTableVisuals(
            ExcelDocument document,
            string sheetName,
            ExcelPdfSaveOptions options) {
            if (!options.UseWorksheetCellStyles) {
                return Array.Empty<StructuredTableVisualData>();
            }

            var visuals = new List<StructuredTableVisualData>();
            foreach (ExcelTableInfo table in document.GetTables()) {
                if (!string.Equals(table.SheetName, sheetName, StringComparison.OrdinalIgnoreCase) ||
                    !A1.TryParseRange(table.Range, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    continue;
                }

                if (!TryCreateStructuredTablePalette(document, table.StyleName, out StructuredTablePalette? palette)) {
                    if (!string.IsNullOrWhiteSpace(table.StyleName)) {
                        AddWarning(
                            options,
                            sheetName,
                            "WorksheetTableStyle",
                            $"Excel table '{table.DisplayName}' uses custom or unknown style '{table.StyleName}'. Its values and direct cell formatting were preserved, but the table style could not be projected.");
                    }
                    continue;
                }

                visuals.Add(new StructuredTableVisualData(
                    firstRow,
                    firstColumn,
                    lastRow,
                    lastColumn,
                    table.HasHeaderRow,
                    table.TotalsRowShown,
                    table.ShowFirstColumn,
                    table.ShowLastColumn,
                    table.ShowRowStripes,
                    table.ShowColumnStripes,
                    palette!));
            }

            return visuals;
        }

        private static StructuredTableCellVisual? GetStructuredTableCellVisual(
            IReadOnlyList<StructuredTableVisualData>? tables,
            string?[,]? cellReferences,
            int row,
            int column) {
            if (tables == null || tables.Count == 0 || cellReferences == null ||
                row < 0 || column < 0 ||
                row >= cellReferences.GetLength(0) || column >= cellReferences.GetLength(1)) {
                return null;
            }

            string? reference = cellReferences[row, column];
            if (string.IsNullOrWhiteSpace(reference)) {
                return null;
            }

            (int Row, int Col) cell = A1.ParseCellRef(reference!.Replace("$", string.Empty));
            if (cell.Row <= 0 || cell.Col <= 0) {
                return null;
            }

            for (int index = tables.Count - 1; index >= 0; index--) {
                if (tables[index].TryResolve(cell.Row, cell.Col, out StructuredTableCellVisual? visual)) {
                    return visual;
                }
            }

            return null;
        }

        private static bool TryCreateStructuredTablePalette(
            ExcelDocument document,
            string? styleName,
            out StructuredTablePalette? palette) {
            palette = null;
            if (!TryParseBuiltInTableStyle(styleName, out string? family, out int index)) {
                return false;
            }

            string light = ResolveThemeRgb(document, 0U, null, "FFFFFF");
            string dark = ResolveThemeRgb(document, 1U, null, "000000");
            int familyIndex = (index - 1) % 7;
            string baseColor = ResolveFamilyBaseColor(document, familyIndex);
            string paleColor = ResolveFamilyTintColor(document, familyIndex, 0.8D);
            string stripeColor = ResolveFamilyTintColor(document, familyIndex, 0.6D);
            string mutedColor = ResolveFamilyTintColor(document, familyIndex, -0.25D);

            switch (family) {
                case "Light":
                    if (index <= 7) {
                        palette = new StructuredTablePalette(
                            headerFill: null,
                            headerText: familyIndex == 0 ? dark : mutedColor,
                            bodyFill: null,
                            stripeFill: paleColor,
                            bodyText: familyIndex == 0 ? dark : mutedColor,
                            border: familyIndex == 0 ? ResolveThemeRgb(document, 0U, -0.35D, "A6A6A6") : baseColor,
                            headerBold: true);
                    } else if (index <= 14) {
                        palette = new StructuredTablePalette(
                            baseColor,
                            light,
                            bodyFill: null,
                            stripeFill: null,
                            dark,
                            baseColor,
                            headerBold: true);
                    } else {
                        palette = new StructuredTablePalette(
                            headerFill: null,
                            headerText: dark,
                            bodyFill: null,
                            stripeFill: paleColor,
                            bodyText: dark,
                            border: familyIndex == 0 ? ResolveThemeRgb(document, 0U, -0.35D, "A6A6A6") : baseColor,
                            headerBold: true);
                    }
                    return true;

                case "Medium":
                    if (index <= 7) {
                        palette = new StructuredTablePalette(baseColor, light, null, paleColor, dark, baseColor, headerBold: true);
                    } else if (index <= 14) {
                        palette = new StructuredTablePalette(baseColor, light, paleColor, stripeColor, dark, baseColor, headerBold: true);
                    } else if (index <= 21) {
                        string neutralStripe = ResolveThemeRgb(document, 0U, -0.15D, "D9D9D9");
                        palette = new StructuredTablePalette(baseColor, light, null, neutralStripe, dark, baseColor, headerBold: true);
                    } else {
                        palette = new StructuredTablePalette(paleColor, dark, paleColor, stripeColor, dark, baseColor, headerBold: true);
                    }
                    return true;

                case "Dark":
                    if (index <= 7) {
                        string bodyFill = familyIndex == 0
                            ? ResolveThemeRgb(document, 1U, 0.45D, "737373")
                            : baseColor;
                        string bodyStripe = familyIndex == 0
                            ? ResolveThemeRgb(document, 1U, 0.25D, "404040")
                            : mutedColor;
                        palette = new StructuredTablePalette(dark, light, bodyFill, bodyStripe, light, baseColor, headerBold: true);
                    } else if (index == 8) {
                        palette = new StructuredTablePalette(
                            dark,
                            light,
                            ResolveThemeRgb(document, 0U, -0.15D, "D9D9D9"),
                            ResolveThemeRgb(document, 0U, -0.35D, "A6A6A6"),
                            dark,
                            ResolveThemeRgb(document, 0U, -0.35D, "A6A6A6"),
                            headerBold: false);
                    } else {
                        uint headerTheme = index == 9 ? 5U : index == 10 ? 7U : 9U;
                        uint bodyTheme = index == 9 ? 4U : index == 10 ? 6U : 8U;
                        string header = ResolveThemeRgb(document, headerTheme, null, index == 9 ? "E97132" : index == 10 ? "0F9ED5" : "4EA72E");
                        string body = ResolveThemeRgb(document, bodyTheme, 0.8D, index == 9 ? "C0E6F5" : index == 10 ? "C1F0C8" : "F2CEEF");
                        string stripe = ResolveThemeRgb(document, bodyTheme, 0.6D, index == 9 ? "83CCEB" : index == 10 ? "83E28E" : "E49EDD");
                        palette = new StructuredTablePalette(header, light, body, stripe, dark, header, headerBold: false);
                    }
                    return true;
            }

            return false;
        }

        private static bool TryParseBuiltInTableStyle(string? styleName, out string? family, out int index) {
            family = null;
            index = 0;
            if (string.IsNullOrWhiteSpace(styleName) ||
                !styleName!.StartsWith("TableStyle", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            string suffix = styleName.Substring("TableStyle".Length);
            foreach (string candidate in new[] { "Light", "Medium", "Dark" }) {
                if (!suffix.StartsWith(candidate, StringComparison.OrdinalIgnoreCase) ||
                    !int.TryParse(suffix.Substring(candidate.Length), out int parsed)) {
                    continue;
                }

                int maximum = candidate == "Light" ? 21 : candidate == "Medium" ? 28 : 11;
                if (parsed < 1 || parsed > maximum) {
                    return false;
                }

                family = candidate;
                index = parsed;
                return true;
            }

            return false;
        }

        private static string ResolveFamilyBaseColor(ExcelDocument document, int familyIndex) {
            if (familyIndex == 0) {
                return ResolveThemeRgb(document, 1U, null, "000000");
            }

            string[] fallbacks = { "156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E" };
            return ResolveThemeRgb(document, (uint)(familyIndex + 3), null, fallbacks[familyIndex - 1]);
        }

        private static string ResolveFamilyTintColor(ExcelDocument document, int familyIndex, double tint) {
            if (familyIndex == 0) {
                string neutralFallback = tint >= 0D
                    ? tint >= 0.7D ? "D9D9D9" : tint >= 0.5D ? "A6A6A6" : "737373"
                    : "A6A6A6";
                double lightTint = tint >= 0D ? tint - 0.95D : tint;
                return ResolveThemeRgb(document, 0U, lightTint, neutralFallback);
            }

            string[] paleFallbacks = { "C0E6F5", "FBE2D5", "C1F0C8", "CAEDFB", "F2CEEF", "DAF2D0" };
            string[] stripeFallbacks = { "83CCEB", "F7C7AC", "83E28E", "94DCF8", "E49EDD", "B5E6A2" };
            string[] mutedFallbacks = { "104861", "BE5014", "12501A", "0C769E", "782170", "3C7D22" };
            string fallback = tint >= 0.7D
                ? paleFallbacks[familyIndex - 1]
                : tint >= 0D
                    ? stripeFallbacks[familyIndex - 1]
                    : mutedFallbacks[familyIndex - 1];
            return ResolveThemeRgb(document, (uint)(familyIndex + 3), tint, fallback);
        }

        private static string ResolveThemeRgb(ExcelDocument document, uint themeIndex, double? tint, string fallback) {
            string? argb = document.ResolveThemeColorArgb(themeIndex, tint);
            if (string.IsNullOrWhiteSpace(argb)) {
                return fallback;
            }

            string normalized = argb!.Trim().TrimStart('#');
            return normalized.Length == 8 ? normalized.Substring(2) : normalized.Length == 6 ? normalized : fallback;
        }

        private sealed class StructuredTableVisualData {
            private readonly int _firstRow;
            private readonly int _firstColumn;
            private readonly int _lastRow;
            private readonly int _lastColumn;
            private readonly bool _hasHeader;
            private readonly bool _hasTotals;
            private readonly bool _showFirstColumn;
            private readonly bool _showLastColumn;
            private readonly bool _showRowStripes;
            private readonly bool _showColumnStripes;
            private readonly StructuredTablePalette _palette;

            public StructuredTableVisualData(
                int firstRow,
                int firstColumn,
                int lastRow,
                int lastColumn,
                bool hasHeader,
                bool hasTotals,
                bool showFirstColumn,
                bool showLastColumn,
                bool showRowStripes,
                bool showColumnStripes,
                StructuredTablePalette palette) {
                _firstRow = firstRow;
                _firstColumn = firstColumn;
                _lastRow = lastRow;
                _lastColumn = lastColumn;
                _hasHeader = hasHeader;
                _hasTotals = hasTotals;
                _showFirstColumn = showFirstColumn;
                _showLastColumn = showLastColumn;
                _showRowStripes = showRowStripes;
                _showColumnStripes = showColumnStripes;
                _palette = palette;
            }

            public bool TryResolve(int row, int column, out StructuredTableCellVisual? visual) {
                visual = null;
                if (row < _firstRow || row > _lastRow || column < _firstColumn || column > _lastColumn) {
                    return false;
                }

                bool header = _hasHeader && row == _firstRow;
                bool totals = _hasTotals && row == _lastRow;
                int bodyRow = row - _firstRow - (_hasHeader ? 1 : 0);
                int tableColumn = column - _firstColumn;
                string? fill = header ? _palette.HeaderFill : _palette.BodyFill;
                if (!header && !totals) {
                    if (_showRowStripes && bodyRow >= 0 && bodyRow % 2 == 0) {
                        fill = _palette.StripeFill ?? fill;
                    }
                    if (_showColumnStripes && tableColumn % 2 == 0) {
                        fill = _palette.StripeFill ?? fill;
                    }
                }

                bool emphasizedColumn =
                    _showFirstColumn && column == _firstColumn ||
                    _showLastColumn && column == _lastColumn;
                visual = new StructuredTableCellVisual(
                    fill,
                    header ? _palette.HeaderText : _palette.BodyText,
                    header ? _palette.HeaderBold : emphasizedColumn,
                    _palette.Border);
                return true;
            }
        }

        private sealed class StructuredTablePalette {
            public StructuredTablePalette(
                string? headerFill,
                string? headerText,
                string? bodyFill,
                string? stripeFill,
                string? bodyText,
                string? border,
                bool headerBold) {
                HeaderFill = headerFill;
                HeaderText = headerText;
                BodyFill = bodyFill;
                StripeFill = stripeFill;
                BodyText = bodyText;
                Border = border;
                HeaderBold = headerBold;
            }

            public string? HeaderFill { get; }
            public string? HeaderText { get; }
            public string? BodyFill { get; }
            public string? StripeFill { get; }
            public string? BodyText { get; }
            public string? Border { get; }
            public bool HeaderBold { get; }
        }

        private sealed class StructuredTableCellVisual {
            public StructuredTableCellVisual(string? fill, string? text, bool bold, string? border) {
                Fill = fill;
                Text = text;
                Bold = bold;
                Border = border;
            }

            public string? Fill { get; }
            public string? Text { get; }
            public bool Bold { get; }
            public string? Border { get; }
        }
    }
}
