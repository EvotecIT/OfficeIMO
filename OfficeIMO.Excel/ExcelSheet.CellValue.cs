using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SixLabors.Fonts;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using SixLaborsColor = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        
        // Core implementation: single source of truth (no locks here)
        private void CellValueCore(int row, int column, object value)
        {
            var cell = GetCell(row, column);
            var (cellValue, dataType) = CoerceForCell(value);
            cell.CellValue = cellValue;
            cell.DataType = dataType;

            // Automatically apply date format for DateTime values
            // Using Excel's built-in date format code 14 (invariant short date)
            if (value is DateTime || value is DateTimeOffset)
            {
                ApplyBuiltInNumberFormat(row, column, 14);  // Built-in format 14 is short date
            }

            // Enable wrap text when value contains new lines so Excel renders multiple lines correctly
            if (value is string s && (s.Contains("\n") || s.Contains("\r")))
            {
                ApplyWrapText(row, column);
            }
        }

        // Compute-only coercion (no OpenXML mutations, except SharedString for now)
        private (CellValue cellValue, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> dataType) CoerceForCell(object value)
        {
            switch (value)
            {
                case null:
                    return (new CellValue(string.Empty), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.String));
                case string s:
                    // TODO: SharedString index should be resolved via planner in parallel scenarios
                    int sharedStringIndex = _excelDocument.GetSharedStringIndex(s);
                    return (new CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString));
                case double d:
                    return (new CellValue(d.ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case float f:
                    return (new CellValue(Convert.ToDouble(f).ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case decimal dec:
                    return (new CellValue(dec.ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case int i:
                    return (new CellValue(((double)i).ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case long l:
                    return (new CellValue(((double)l).ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case DateTime dt:
                    return (new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case DateTimeOffset dto:
                    return (new CellValue(dto.UtcDateTime.ToOADate().ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
#if NET6_0_OR_GREATER
                case DateOnly dateOnly:
                    return (new CellValue(dateOnly.ToDateTime(TimeOnly.MinValue).ToOADate().ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case TimeOnly timeOnly:
                    return (new CellValue(timeOnly.ToTimeSpan().TotalDays.ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
#endif
                case TimeSpan ts:
                    return (new CellValue(ts.TotalDays.ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case bool b:
                    return (new CellValue(b ? "1" : "0"), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean));
                case uint ui:
                    return (new CellValue(((double)ui).ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case ulong ul:
                    return (new CellValue(((double)ul).ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case ushort us:
                    return (new CellValue(((double)us).ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case byte by:
                    return (new CellValue(((double)by).ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case sbyte sb:
                    return (new CellValue(((double)sb).ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case short sh:
                    return (new CellValue(((double)sh).ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                case Guid guid:
                    {
                        int idx = _excelDocument.GetSharedStringIndex(guid.ToString());
                        return (new CellValue(idx.ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString));
                    }
                case Enum e:
                    {
                        string name = e.ToString();
                        int idx = _excelDocument.GetSharedStringIndex(name);
                        return (new CellValue(idx.ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString));
                    }
                case char ch:
                    {
                        int idx = _excelDocument.GetSharedStringIndex(ch.ToString());
                        return (new CellValue(idx.ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString));
                    }
                case System.DBNull:
                    return (new CellValue(string.Empty), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.String));
                case Uri uri:
                    {
                        int idx = _excelDocument.GetSharedStringIndex(uri.ToString());
                        return (new CellValue(idx.ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString));
                    }
                default:
                    string stringValue = value?.ToString() ?? string.Empty;
                    int defaultIndex = _excelDocument.GetSharedStringIndex(stringValue);
                    return (new CellValue(defaultIndex.ToString(CultureInfo.InvariantCulture)), new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString));
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, string value) {
            WriteLockConditional(() => CellValueCore(row, column, value));
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, double value) {
            WriteLockConditional(() => CellValueCore(row, column, value));
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, decimal value) {
            WriteLockConditional(() => CellValueCore(row, column, value));
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, DateTime value) {
            WriteLockConditional(() => {
                CellValueCore(row, column, value);
                // DateTime formatting is now handled in CellValueCore
            });
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, DateTimeOffset value) {
            CellValue(row, column, value.UtcDateTime);
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, TimeSpan value) {
            WriteLockConditional(() => CellValueCore(row, column, value));
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, uint value) {
            CellValue(row, column, (double)value);
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, ulong value) {
            CellValue(row, column, (double)value);
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, ushort value) {
            CellValue(row, column, (double)value);
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, byte value) {
            CellValue(row, column, (double)value);
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, sbyte value) {
            CellValue(row, column, (double)value);
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, bool value) {
            WriteLockConditional(() => CellValueCore(row, column, value));
        }

        /// <summary>
        /// Sets a formula in the specified cell.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="formula">The formula expression.</param>
        public void CellFormula(int row, int column, string formula) {
            WriteLock(() => {
                Cell cell = GetCell(row, column);
                cell.CellFormula = new CellFormula(formula);
            });
        }

        /// <summary>
        /// Applies bold font to a single cell.
        /// </summary>
        public void CellBold(int row, int column, bool bold = true)
        {
            WriteLockConditional(() =>
            {
                var cell = GetCell(row, column);
                ApplyFontBold(cell, bold);
            });
        }

        /// <summary>
        /// Applies solid background to a single cell. Accepts #RRGGBB or #AARRGGBB.
        /// </summary>
        public void CellBackground(int row, int column, string hexColor)
        {
            if (string.IsNullOrWhiteSpace(hexColor)) return;
            WriteLockConditional(() =>
            {
                var cell = GetCell(row, column);
                ApplyBackground(cell, hexColor);
            });
        }

        /// <summary>
        /// Applies solid background to a single cell using SixLabors color.
        /// </summary>
        public void CellBackground(int row, int column, SixLabors.ImageSharp.Color color)
        {
            var argb = OfficeIMO.Excel.ExcelColor.ToArgbHex(color);
            CellBackground(row, column, argb);
        }

        /// <summary>
        /// Sets the value, formula, and number format of a cell in a single operation.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="value">Optional value to assign.</param>
        /// <param name="formula">Optional formula to apply.</param>
        /// <param name="numberFormat">Optional number format code.</param>
        public void Cell(int row, int column, object? value = null, string? formula = null, string? numberFormat = null) {
            if (value != null) {
                CellValue(row, column, value);
            }
            if (!string.IsNullOrEmpty(formula)) {
                CellFormula(row, column, formula!);
            }
            if (!string.IsNullOrEmpty(numberFormat)) {
                FormatCell(row, column, numberFormat!);
            }
        }

        /// <summary>
        /// Applies a number format to the specified cell.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="numberFormat">The number format code to apply.</param>
        public void FormatCell(int row, int column, string numberFormat) {
            WriteLockConditional(() => FormatCellCore(row, column, numberFormat));
        }

        /// <summary>
        /// Tries to read the display text of a cell at the given position.
        /// Returns false if the cell is blank or out of bounds.
        /// </summary>
        public bool TryGetCellText(int row, int column, out string text)
        {
            text = string.Empty;
            try
            {
                var cell = GetCell(row, column);
                if (cell == null) return false;
                // Resolve shared string if needed
                if (cell.DataType != null && cell.DataType.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString)
                {
                    if (int.TryParse(cell.InnerText, System.Globalization.NumberStyles.Integer, CultureInfo.InvariantCulture, out int ssid))
                    {
                        var wb = _excelDocument._spreadSheetDocument.WorkbookPart;
                        var sst = wb?.SharedStringTablePart?.SharedStringTable;
                        var si = sst?.Elements<SharedStringItem>().ElementAtOrDefault(ssid);
                        if (si != null)
                        {
                            text = si.InnerText ?? string.Empty;
                            return true;
                        }
                        return false;
                    }
                }
                // Otherwise, return inner text (numbers/booleans as invariant string)
                text = cell.InnerText ?? string.Empty;
                return !string.IsNullOrEmpty(text);
            }
            catch { return false; }
        }

        private void ApplyWrapText(int row, int column)
        {
            var cell = GetCell(row, column);
            ApplyWrapText(cell);
        }

        private void ApplyWrapText(Cell cell)
        {
            var workbookPart = _excelDocument._spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
            WorkbookStylesPart? stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null)
            {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            // Base on existing cell's style if present
            uint baseIndex = cell.StyleIndex?.Value ?? 0U;
            var cellFormatsEl = stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            var cellFormats = cellFormatsEl.Elements<CellFormat>().ToList();
            var baseFormat = cellFormats.ElementAtOrDefault((int)baseIndex) ?? new CellFormat
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U
            };

            // Try to find an existing format with same base ids and WrapText enabled
            int wrapIndex = -1;
            for (int i = 0; i < cellFormats.Count; i++)
            {
                var cf = cellFormats[i];
                var align = cf.Alignment;
                bool wrap = align != null && align.WrapText != null && align.WrapText.Value;
                if (wrap && cf.NumberFormatId?.Value == baseFormat.NumberFormatId?.Value
                        && cf.FontId?.Value == baseFormat.FontId?.Value
                        && cf.FillId?.Value == baseFormat.FillId?.Value
                        && cf.BorderId?.Value == baseFormat.BorderId?.Value)
                {
                    wrapIndex = i;
                    break;
                }
            }

            if (wrapIndex == -1)
            {
                var newFormat = new CellFormat
                {
                    NumberFormatId = baseFormat.NumberFormatId ?? 0U,
                    FontId = baseFormat.FontId ?? 0U,
                    FillId = baseFormat.FillId ?? 0U,
                    BorderId = baseFormat.BorderId ?? 0U,
                    FormatId = baseFormat.FormatId ?? 0U,
                    ApplyAlignment = true,
                    Alignment = new Alignment { WrapText = true }
                };
                cellFormatsEl.Append(newFormat);
                cellFormatsEl.Count = (uint)cellFormatsEl.Count();
                wrapIndex = (int)cellFormatsEl.Count.Value - 1;
                stylesPart.Stylesheet.Save();
            }

            cell.StyleIndex = (uint)wrapIndex;
        }

        /// <summary>
        /// Applies a horizontal alignment to a single cell.
        /// </summary>
        public void CellAlign(int row, int column, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues alignment)
        {
            WriteLockConditional(() =>
            {
                var cell = GetCell(row, column);
                var workbookPart = _excelDocument._spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                uint baseIndex = cell.StyleIndex?.Value ?? 0U;
                var cfEl = stylesheet.CellFormats ??= new CellFormats(new CellFormat());
                var cfs = cfEl.Elements<CellFormat>().ToList();
                var baseFormat = cfs.ElementAtOrDefault((int)baseIndex) ?? new CellFormat
                {
                    NumberFormatId = 0U,
                    FontId = 0U,
                    FillId = 0U,
                    BorderId = 0U,
                    FormatId = 0U
                };

                // Try to reuse an existing format with same base ids and requested alignment
                int found = -1;
                for (int i = 0; i < cfs.Count; i++)
                {
                    var cf = cfs[i];
                    var a = cf.Alignment;
                    if (a != null && a.Horizontal != null && a.Horizontal.Value == alignment
                        && cf.NumberFormatId?.Value == baseFormat.NumberFormatId?.Value
                        && cf.FontId?.Value == baseFormat.FontId?.Value
                        && cf.FillId?.Value == baseFormat.FillId?.Value
                        && cf.BorderId?.Value == baseFormat.BorderId?.Value)
                    {
                        found = i; break;
                    }
                }
                if (found == -1)
                {
                    var newFormat = new CellFormat
                    {
                        NumberFormatId = baseFormat.NumberFormatId ?? 0U,
                        FontId = baseFormat.FontId ?? 0U,
                        FillId = baseFormat.FillId ?? 0U,
                        BorderId = baseFormat.BorderId ?? 0U,
                        FormatId = baseFormat.FormatId ?? 0U,
                        ApplyAlignment = true,
                        Alignment = new Alignment { Horizontal = alignment }
                    };
                    cfEl.Append(newFormat);
                    cfEl.Count = (uint)cfEl.Count();
                    stylesPart.Stylesheet.Save();
                    found = (int)cfEl.Count.Value - 1;
                }
                cell.StyleIndex = (uint)found;
            });
        }

        /// <summary>
        /// Applies a font color (ARGB hex or #RRGGBB) to a single cell.
        /// </summary>
        public void CellFontColor(int row, int column, string hexColor)
        {
            if (string.IsNullOrWhiteSpace(hexColor)) return;
            WriteLockConditional(() =>
            {
                var cell = GetCell(row, column);
                var workbookPart = _excelDocument._spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                string argb = NormalizeHexColor(hexColor);

                // Ensure a Font with this color exists
                var fontsEl = stylesheet.Fonts ??= new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font());
                var fonts = fontsEl.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ToList();
                int fontId = -1;
                for (int i = 0; i < fonts.Count; i++)
                {
                    var c = fonts[i].Color?.Rgb?.Value;
                    var b = fonts[i].Bold != null ? true : false;
                    if (c == argb && !b) { fontId = i; break; }
                }
                if (fontId == -1)
                {
                var f = new DocumentFormat.OpenXml.Spreadsheet.Font();
                f.Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = argb };
                fontsEl.Append(f);
                fontsEl.Count = (uint)fontsEl.Count();
                fontId = (int)fontsEl.Count.Value - 1;
                }

                uint baseIndex = cell.StyleIndex?.Value ?? 0U;
                var cellFormatsEl2 = stylesheet.CellFormats ??= new CellFormats(new CellFormat());
                var cellFormats = cellFormatsEl2.Elements<CellFormat>().ToList();
                var baseFormat = cellFormats.ElementAtOrDefault((int)baseIndex) ?? new CellFormat
                {
                    NumberFormatId = 0U,
                    FontId = 0U,
                    FillId = 0U,
                    BorderId = 0U,
                    FormatId = 0U,
                };
                var newFormat2 = new CellFormat
                {
                    NumberFormatId = baseFormat.NumberFormatId ?? 0U,
                    FontId = (uint)fontId,
                    FillId = baseFormat.FillId ?? 0U,
                    BorderId = baseFormat.BorderId ?? 0U,
                    FormatId = baseFormat.FormatId ?? 0U,
                };
                cellFormatsEl2.Append(newFormat2);
                cellFormatsEl2.Count = (uint)cellFormatsEl2.Count();
                cell.StyleIndex = (uint)cellFormatsEl2.Count.Value - 1;
                stylesPart.Stylesheet.Save();
            });
        }

        private void ApplyFontBold(Cell cell, bool bold)
        {
            var workbookPart = _excelDocument._spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null)
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            // Ensure we have a bold font entry
            int boldFontId = -1;
            var fontsEl = stylesheet.Fonts ??= new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font());
            var fonts = fontsEl.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ToList();
            for (int i = 0; i < fonts.Count; i++)
            {
                bool hasBold = fonts[i].Bold != null;
                if (hasBold == bold)
                {
                    boldFontId = i;
                    break;
                }
            }
            if (boldFontId == -1)
            {
                var f = new DocumentFormat.OpenXml.Spreadsheet.Font();
                if (bold) f.Bold = new Bold();
                fontsEl.Append(f);
                fontsEl.Count = (uint)fontsEl.Count();
                boldFontId = (int)fontsEl.Count.Value - 1;
            }

            uint baseIndex = cell.StyleIndex?.Value ?? 0U;
            var cellFormatsEl3 = stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            var cellFormats = cellFormatsEl3.Elements<CellFormat>().ToList();
            var baseFormat = cellFormats.ElementAtOrDefault((int)baseIndex) ?? new CellFormat
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
            };

            // Create a new CellFormat using the bold font, preserving other IDs
            var newFormat = new CellFormat
            {
                NumberFormatId = baseFormat.NumberFormatId ?? 0U,
                FontId = (uint)boldFontId,
                FillId = baseFormat.FillId ?? 0U,
                BorderId = baseFormat.BorderId ?? 0U,
                FormatId = baseFormat.FormatId ?? 0U,
                ApplyFont = true
            };
            cellFormatsEl3.Append(newFormat);
            cellFormatsEl3.Count = (uint)cellFormatsEl3.Count();
            cell.StyleIndex = (uint)cellFormatsEl3.Count.Value - 1;
            stylesPart.Stylesheet.Save();
        }

        private static string NormalizeHexColor(string hex)
        {
            hex = hex.Trim();
            if (hex.StartsWith("#")) hex = hex.Substring(1);
            if (hex.Length == 6) return "FF" + hex.ToUpperInvariant();
            if (hex.Length == 8) return hex.ToUpperInvariant();
            // Fallback default
            return "FFFFFFFF";
        }

        private void ApplyBackground(Cell cell, string hexColor)
        {
            var workbookPart = _excelDocument._spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null)
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            // Create a fill with solid color
            string argb = NormalizeHexColor(hexColor);
            var fill = new Fill(new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = argb },
                BackgroundColor = new BackgroundColor { Rgb = argb }
            });
            var fillsEl = stylesheet.Fills ??= new Fills();
            fillsEl.Append(fill);
            fillsEl.Count = (uint)fillsEl.Count();
            int fillId = (int)fillsEl.Count.Value - 1;

            uint baseIndex = cell.StyleIndex?.Value ?? 0U;
            var cellFormatsEl = stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            var cellFormats = cellFormatsEl.Elements<CellFormat>().ToList();
            var baseFormat = cellFormats.ElementAtOrDefault((int)baseIndex) ?? new CellFormat
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
            };

            var newFormat = new CellFormat
            {
                NumberFormatId = baseFormat.NumberFormatId ?? 0U,
                FontId = baseFormat.FontId ?? 0U,
                FillId = (uint)fillId,
                BorderId = baseFormat.BorderId ?? 0U,
                FormatId = baseFormat.FormatId ?? 0U,
                ApplyFill = true
            };
            stylesheet.CellFormats.Append(newFormat);
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
            cell.StyleIndex = (uint)stylesheet.CellFormats.Count.Value - 1;
            stylesPart.Stylesheet.Save();
        }

        private void ApplyBuiltInNumberFormat(int row, int column, uint builtInFormatId) {
            Cell cell = GetCell(row, column);

            var workbookPart = _excelDocument._spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
            WorkbookStylesPart? stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            var cellFormatsEl = stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            var cellFormats = cellFormatsEl.Elements<CellFormat>().ToList();
            int formatIndex = cellFormats.FindIndex(cf => cf.NumberFormatId != null && cf.NumberFormatId.Value == builtInFormatId && cf.ApplyNumberFormat != null && cf.ApplyNumberFormat.Value);
            if (formatIndex == -1) {
                CellFormat cellFormat = new CellFormat {
                    NumberFormatId = builtInFormatId,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,
                    FormatId = 0,
                    ApplyNumberFormat = true
                };
                stylesheet.CellFormats.Append(cellFormat);
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                formatIndex = cellFormats.Count;
            }

            cell.StyleIndex = (uint)formatIndex;
            stylesPart.Stylesheet.Save();
        }

        private void FormatCellCore(int row, int column, string numberFormat) {
            Cell cell = GetCell(row, column);

            var workbookPart = _excelDocument._spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
            WorkbookStylesPart? stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            stylesheet.NumberingFormats ??= new NumberingFormats();
            NumberingFormat? existingFormat = stylesheet.NumberingFormats.Elements<NumberingFormat>()
                .FirstOrDefault(n => n.FormatCode != null && n.FormatCode.Value == numberFormat);

            uint numberFormatId;
            if (existingFormat != null) {
                numberFormatId = existingFormat.NumberFormatId!.Value;
            } else {
                numberFormatId = stylesheet.NumberingFormats.Elements<NumberingFormat>().Any()
                    ? stylesheet.NumberingFormats.Elements<NumberingFormat>().Max(n => n.NumberFormatId!.Value) + 1
                    : 164U;
                NumberingFormat numberingFormat = new NumberingFormat {
                    NumberFormatId = numberFormatId,
                    FormatCode = StringValue.FromString(numberFormat)
                };
                stylesheet.NumberingFormats.Append(numberingFormat);
                stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Count();
            }

            var cellFormatsEl4 = stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            var cellFormats = cellFormatsEl4.Elements<CellFormat>().ToList();
            int formatIndex = cellFormats.FindIndex(cf => cf.NumberFormatId != null && cf.NumberFormatId.Value == numberFormatId && cf.ApplyNumberFormat != null && cf.ApplyNumberFormat.Value);
            if (formatIndex == -1) {
                CellFormat cellFormat = new CellFormat {
                    NumberFormatId = numberFormatId,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,
                    FormatId = 0,
                    ApplyNumberFormat = true
                };
                cellFormatsEl4.Append(cellFormat);
                cellFormatsEl4.Count = (uint)cellFormatsEl4.Count();
                formatIndex = cellFormats.Count;
            }

            cell.StyleIndex = (uint)formatIndex;
            stylesPart.Stylesheet.Save();
        }

        /// <summary>
        /// Ensures required default style primitives exist and their counts are consistent.
        /// Excel expects at least 1 Font, 2 Fills (None, Gray125), 1 Border,
        /// 1 CellStyleFormat, and 1 CellFormat present.
        /// </summary>
        private static void EnsureDefaultStylePrimitives(Stylesheet stylesheet)
        {
            // Fonts
            if (stylesheet.Fonts == null || !stylesheet.Fonts.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().Any())
            {
                stylesheet.Fonts = new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font());
            }
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

            // Fills: ensure index 0 = None, index 1 = Gray125
            if (stylesheet.Fills == null)
            {
                stylesheet.Fills = new Fills();
            }
            var fills = stylesheet.Fills.Elements<Fill>().ToList();
            bool hasNone = fills.Any(f => f.PatternFill?.PatternType?.Value == PatternValues.None);
            bool hasGray = fills.Any(f => f.PatternFill?.PatternType?.Value == PatternValues.Gray125);
            if (!hasNone)
            {
                stylesheet.Fills.AppendChild(new Fill(new PatternFill { PatternType = PatternValues.None }));
            }
            if (!hasGray)
            {
                stylesheet.Fills.AppendChild(new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
            }
            stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();

            // Borders
            if (stylesheet.Borders == null || !stylesheet.Borders.Elements<Border>().Any())
            {
                stylesheet.Borders = new Borders(new Border());
            }
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();

            // Cell style formats
            if (stylesheet.CellStyleFormats == null || !stylesheet.CellStyleFormats.Elements<CellFormat>().Any())
            {
                stylesheet.CellStyleFormats = new CellStyleFormats(new CellFormat());
            }
            stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.Count();

            // Cell formats
            if (stylesheet.CellFormats == null || !stylesheet.CellFormats.Elements<CellFormat>().Any())
            {
                stylesheet.CellFormats = new CellFormats(new CellFormat());
            }
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();

            // Numbering formats count normalization
            if (stylesheet.NumberingFormats != null)
            {
                stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Count();
            }
        }

        /// <summary>
        /// Sets the specified value into a cell, inferring the data type.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="value">The value to assign.</param>
        public void CellValue(int row, int column, object value) {
            WriteLockConditional(() => CellValueCore(row, column, value));
        }

        /// <summary>
        /// Sets the value of a cell using a nullable struct.
        /// </summary>
        /// <typeparam name="T">The value type.</typeparam>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="value">The nullable value to assign.</param>
        public void CellValue<T>(int row, int column, T? value) where T : struct {
            if (value.HasValue) {
                CellValue(row, column, value.Value);
            } else {
                CellValue(row, column, string.Empty);
            }
        }

    }
}

