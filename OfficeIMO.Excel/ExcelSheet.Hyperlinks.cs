using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static string EscapeSheetNameForLink(string name) {
            return (name ?? string.Empty).Replace("'", "''");
        }
        /// <summary>
        /// Sets an external hyperlink on a single cell. If <paramref name="display"/> is null or empty, the URL is shown.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="url">Target URL.</param>
        /// <param name="display">Optional display text. When null/empty, the URL is displayed.</param>
        /// <param name="style">When true, applies hyperlink styling (blue + underline).</param>
        public void SetHyperlink(int row, int column, string url, string? display = null, bool style = true) {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentNullException(nameof(url));

            WriteLock(() => {
                var cell = GetCell(row, column);
                string text = string.IsNullOrEmpty(display) ? url : display!;
                // Avoid nested locks: write value using core method
                CellValueCore(row, column, text);

                var reference = GetColumnName(column) + row.ToString(System.Globalization.CultureInfo.InvariantCulture);
                // Ensure Hyperlinks container exists
                var ws = _worksheetPart.Worksheet;
                var hyperlinks = ws.Elements<Hyperlinks>().FirstOrDefault();
                if (hyperlinks == null) {
                    hyperlinks = new Hyperlinks();
                    // place near the end but before TableParts per schema order
                    var tableParts = ws.Elements<TableParts>().FirstOrDefault();
                    if (tableParts != null) ws.InsertBefore(hyperlinks, tableParts); else ws.Append(hyperlinks);
                } else {
                    RemoveHyperlinksByReference(hyperlinks, reference);
                }

                // Add external relationship
                var rel = _worksheetPart.AddHyperlinkRelationship(new Uri(url), true);
                var hl = new Hyperlink { Reference = reference, Id = rel.Id };
                hyperlinks.Append(hl);
                if (style) ApplyHyperlinkStyle(cell);
            });
        }
        /// <example>
        /// sheet.SetHyperlink(2, 1, "https://example.org", display: "Example", style: true);
        /// </example>

        /// <summary>
        /// Sets an external hyperlink using an A1 reference (e.g., "B5").
        /// </summary>
        /// <param name="a1">A1 cell reference without a sheet prefix.</param>
        /// <param name="url">Target URL.</param>
        /// <param name="display">Optional display text. When null/empty, the URL is displayed.</param>
        /// <param name="style">When true, applies hyperlink styling (blue + underline).</param>
        public void SetHyperlink(string a1, string url, string? display = null, bool style = true) {
            var col = GetColumnIndex(a1);
            var row = GetRowIndex(a1);
            SetHyperlink(row, col, url, display, style);
        }
        /// <example>
        /// sheet.SetHyperlink("B5", "https://contoso.com");
        /// </example>

        /// <summary>
        /// Sets an internal hyperlink (location in this workbook), e.g., "'Sheet1'!A1".
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="location">Target location inside the workbook (e.g., "'Summary'!A1").</param>
        /// <param name="display">Optional display text. When null/empty, <paramref name="location"/> is displayed.</param>
        /// <param name="style">When true, applies hyperlink styling (blue + underline).</param>
        public void SetInternalLink(int row, int column, string location, string? display = null, bool style = true) {
            if (string.IsNullOrWhiteSpace(location)) throw new ArgumentNullException(nameof(location));
            WriteLock(() => {
                string text = string.IsNullOrEmpty(display) ? location : display!;
                CellValueCore(row, column, text);
                var reference = GetColumnName(column) + row.ToString(System.Globalization.CultureInfo.InvariantCulture);
                var ws = _worksheetPart.Worksheet;
                var hyperlinks = ws.Elements<Hyperlinks>().FirstOrDefault();
                if (hyperlinks == null) {
                    hyperlinks = new Hyperlinks();
                    var tableParts = ws.Elements<TableParts>().FirstOrDefault();
                    if (tableParts != null) ws.InsertBefore(hyperlinks, tableParts); else ws.Append(hyperlinks);
                } else {
                    RemoveHyperlinksByReference(hyperlinks, reference);
                }
                var hl = new Hyperlink { Reference = reference, Location = location };
                hyperlinks.Append(hl);
                // Defer save to caller; final document Save() will persist
                var cell = GetCell(row, column);
                if (style) ApplyHyperlinkStyle(cell);
            });
        }
        /// <example>
        /// // Link A2 to the top of Summary sheet
        /// sheet.SetInternalLink(2, 1, "'Summary'!A1", display: "Summary", style: true);
        /// </example>

        /// <summary>
        /// Sets an internal hyperlink to a target sheet and A1 location using safe quoting rules.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="target">Target sheet.</param>
        /// <param name="a1">A1 reference on the target sheet.</param>
        /// <param name="display">Optional display text. When null/empty, the location is displayed.</param>
        /// <param name="style">When true, applies hyperlink styling (blue + underline).</param>
        public void SetInternalLink(int row, int column, ExcelSheet target, string a1, string? display = null, bool style = true) {
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (string.IsNullOrWhiteSpace(a1)) throw new ArgumentNullException(nameof(a1));
            string loc = $"'{EscapeSheetNameForLink(target.Name)}'!{a1}";
            SetInternalLink(row, column, loc, display, style);
        }

        private void ApplyHyperlinkStyle(Cell cell) {
            var workbookPart = _excelDocument._spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            // Ensure primitives
            EnsureDefaultStylePrimitives(stylesheet);

            // Ensure hyperlink font (blue + underline) exists
            const string hyperlinkRgb = "FF0563C1"; // Excel default hyperlink blue
            var fontsEl = stylesheet.Fonts ??= new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font());
            int hyperlinkFontIndex = -1;
            int idx = 0;
            foreach (var f in fontsEl.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>()) {
                bool hasUnderline = f.Underline != null;
                string rgb = f.Color?.Rgb?.Value ?? string.Empty;
                if (hasUnderline && string.Equals(rgb, hyperlinkRgb, StringComparison.OrdinalIgnoreCase)) {
                    hyperlinkFontIndex = idx;
                    break;
                }
                idx++;
            }
            if (hyperlinkFontIndex == -1) {
                var font = new DocumentFormat.OpenXml.Spreadsheet.Font {
                    Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = hyperlinkRgb },
                    Underline = new DocumentFormat.OpenXml.Spreadsheet.Underline()
                };
                fontsEl.Append(font);
                hyperlinkFontIndex = fontsEl.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().Count() - 1;
            }

            // Base on existing style
            uint baseIndex = cell.StyleIndex?.Value ?? 0U;
            var cellFormatsEl = stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            var cellFormats = cellFormatsEl.Elements<CellFormat>().ToList();
            var baseFormat = cellFormats.ElementAtOrDefault((int)baseIndex) ?? new CellFormat {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U
            };

            // Try to find an existing format matching base but with hyperlink font
            int hyperlinkFormatIndex = -1;
            for (int i = 0; i < cellFormats.Count; i++) {
                var cf = cellFormats[i];
                if ((cf.NumberFormatId?.Value ?? 0U) == (baseFormat.NumberFormatId?.Value ?? 0U)
                    && (cf.FillId?.Value ?? 0U) == (baseFormat.FillId?.Value ?? 0U)
                    && (cf.BorderId?.Value ?? 0U) == (baseFormat.BorderId?.Value ?? 0U)
                    && (cf.FormatId?.Value ?? 0U) == (baseFormat.FormatId?.Value ?? 0U)
                    && (cf.FontId?.Value ?? 0U) == (uint)hyperlinkFontIndex) {
                    hyperlinkFormatIndex = i;
                    break;
                }
            }
            if (hyperlinkFormatIndex == -1) {
                var newFormat = new CellFormat {
                    NumberFormatId = baseFormat.NumberFormatId ?? 0U,
                    FontId = (uint)hyperlinkFontIndex,
                    FillId = baseFormat.FillId ?? 0U,
                    BorderId = baseFormat.BorderId ?? 0U,
                    FormatId = baseFormat.FormatId ?? 0U,
                    ApplyFont = true
                };
                cellFormatsEl.Append(newFormat);
                hyperlinkFormatIndex = cellFormatsEl.Elements<CellFormat>().Count() - 1;
                stylesPart.Stylesheet.Save();
            }

            cell.StyleIndex = (uint)hyperlinkFormatIndex;
        }

        private void RemoveHyperlinksByReference(Hyperlinks hyperlinks, string reference) {
            var matches = hyperlinks.Elements<Hyperlink>()
                .Where(h => string.Equals(h.Reference?.Value, reference, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (matches.Count == 0) return;

            foreach (var existing in matches) {
                var relId = existing.Id?.Value;
                if (!string.IsNullOrEmpty(relId)) {
                    var rel = _worksheetPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
                    if (rel != null) {
                        _worksheetPart.DeleteReferenceRelationship(rel);
                    }
                }
                existing.Remove();
            }
        }
    }
}
