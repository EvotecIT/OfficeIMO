using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        /// <summary>
        /// Applies bold font to a single cell.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to modify.</param>
        /// <param name="column">The 1-based column index of the cell to modify.</param>
        /// <param name="bold">Whether the font should be bold (true) or regular (false).</param>
        public void CellBold(int row, int column, bool bold = true) {
            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                ApplyFontBold(cell, bold);
            });
        }

        /// <summary>
        /// Applies italic font styling to a single cell.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to modify.</param>
        /// <param name="column">The 1-based column index of the cell to modify.</param>
        /// <param name="italic">Whether the font should be italic (true) or regular (false).</param>
        public void CellItalic(int row, int column, bool italic = true) {
            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                ApplyFontItalic(cell, italic);
            });
        }

        /// <summary>
        /// Applies underline font styling to a single cell.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to modify.</param>
        /// <param name="column">The 1-based column index of the cell to modify.</param>
        /// <param name="underline">Whether the font should be underlined (true) or not (false).</param>
        public void CellUnderline(int row, int column, bool underline = true) {
            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                ApplyFontUnderline(cell, underline);
            });
        }

        /// <summary>
        /// Applies or clears wrap text on a single cell.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to modify.</param>
        /// <param name="column">The 1-based column index of the cell to modify.</param>
        /// <param name="wrapText">Whether text should wrap within the cell.</param>
        public void CellWrapText(int row, int column, bool wrapText = true) {
            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                ApplyCellFormatOverride(stylesheet, cell, format => {
                    var alignment = format.Alignment != null
                        ? (Alignment)format.Alignment.CloneNode(true)
                        : new Alignment();
                    alignment.WrapText = wrapText ? true : null;
                    format.Alignment = alignment;
                    format.ApplyAlignment = true;
                });

                stylesPart.Stylesheet.Save();
            });
        }

        /// <summary>
        /// Applies a font family name to a single cell.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to modify.</param>
        /// <param name="column">The 1-based column index of the cell to modify.</param>
        /// <param name="fontName">The font family name to assign.</param>
        public void CellFontName(int row, int column, string fontName) {
            if (string.IsNullOrWhiteSpace(fontName)) return;
            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                ApplyFontName(cell, fontName);
            });
        }

        /// <summary>
        /// Applies a font size to a single cell.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to modify.</param>
        /// <param name="column">The 1-based column index of the cell to modify.</param>
        /// <param name="fontSize">The font size in points.</param>
        public void CellFontSize(int row, int column, double fontSize) {
            if (fontSize <= 0D || double.IsNaN(fontSize) || double.IsInfinity(fontSize)) {
                throw new ArgumentOutOfRangeException(nameof(fontSize), "Font size must be a positive finite value.");
            }

            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                ApplyFontSize(cell, fontSize);
            });
        }

        /// <summary>
        /// Applies solid background to a single cell. Accepts #RRGGBB or #AARRGGBB.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to fill.</param>
        /// <param name="column">The 1-based column index of the cell to fill.</param>
        /// <param name="hexColor">The background color expressed as an ARGB or RGB hex string.</param>
        public void CellBackground(int row, int column, string hexColor) {
            if (string.IsNullOrWhiteSpace(hexColor)) return;
            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                ApplyBackground(cell, hexColor);
            });
        }

        /// <summary>
        /// Applies solid background to a single cell using an OfficeIMO color.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to fill.</param>
        /// <param name="column">The 1-based column index of the cell to fill.</param>
        /// <param name="color">The <see cref="OfficeIMO.Drawing.OfficeColor"/> to convert to a hex value.</param>
        public void CellBackground(int row, int column, OfficeIMO.Drawing.OfficeColor color) {
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
            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLockConditional(() => FormatCellCore(row, column, numberFormat));
        }

        /// <summary>
        /// Tries to read the display text of a cell at the given position.
        /// Returns false if the cell is blank or out of bounds.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to inspect.</param>
        /// <param name="column">The 1-based column index of the cell to inspect.</param>
        /// <param name="text">When this method returns, contains the extracted cell text if successful; otherwise, an empty string.</param>
        /// <returns><see langword="true"/> if text was read successfully; otherwise, <see langword="false"/>.</returns>
        public bool TryGetCellText(int row, int column, out string text) {
            text = string.Empty;
            try {
                if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                    if (_excelDocument.HasDeferredDirectDataSetImport) {
                        _excelDocument.MaterializeDeferredDataSetImportPreservingFastSaveModel();
                    } else {
                        MaterializeDeferredDataSetImportIfNeeded();
                    }
                }

                var cell = TryGetCell(row, column);
                if (cell == null) return false;
                // Resolve shared string if needed
                if (cell.DataType != null && cell.DataType.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                    if (TryParseCellTextSharedStringIndex(cell.InnerText, out int ssid)) {
                        string? sharedText = BuildCellTextSharedStringSnapshot().Get(ssid);
                        if (sharedText != null) {
                            text = sharedText;
                            return true;
                        }

                        return false;
                    }
                }
                text = GetCellText(cell);
                if (string.IsNullOrEmpty(text) && cell.CellFormula != null && cell.CellValue == null && cell.InlineString == null) {
                    text = cell.CellFormula.Text ?? string.Empty;
                }

                return cell.CellValue != null || cell.InlineString != null || !string.IsNullOrEmpty(text);
            } catch { return false; }
        }

        /// <summary>
        /// Tries to read the native value metadata of a cell at the given position.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to inspect.</param>
        /// <param name="column">The 1-based column index of the cell to inspect.</param>
        /// <param name="snapshot">When this method returns, contains the cell value metadata if successful; otherwise, <see langword="null"/>.</param>
        /// <returns><see langword="true"/> when the cell exists and carries a value, inline string, or formula.</returns>
        public bool TryGetCellValueSnapshot(int row, int column, out ExcelCellValueSnapshot? snapshot) {
            snapshot = null;
            try {
                if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                    if (_excelDocument.HasDeferredDirectDataSetImport) {
                        _excelDocument.MaterializeDeferredDataSetImportPreservingFastSaveModel();
                    } else {
                        MaterializeDeferredDataSetImportIfNeeded();
                    }
                }

                Cell? cell = TryGetCell(row, column);
                if (cell == null) {
                    return false;
                }

                string text = GetCellText(cell);
                if (string.IsNullOrEmpty(text) && cell.CellFormula != null && cell.CellValue == null && cell.InlineString == null) {
                    text = cell.CellFormula.Text ?? string.Empty;
                }

                if (cell.CellValue == null && cell.InlineString == null && string.IsNullOrEmpty(text)) {
                    return false;
                }

                DocumentFormat.OpenXml.Spreadsheet.CellValues? dataType = cell.DataType?.Value;
                ExcelCellValueKind kind;
                if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString ||
                    dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.String ||
                    dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                    kind = ExcelCellValueKind.Text;
                } else if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number) {
                    kind = ExcelCellValueKind.Number;
                } else if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean) {
                    kind = ExcelCellValueKind.Boolean;
                } else if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error) {
                    kind = ExcelCellValueKind.Error;
                } else if (cell.CellFormula != null) {
                    kind = ExcelCellValueKind.Formula;
                } else {
                    kind = dataType == null ? ExcelCellValueKind.Number : ExcelCellValueKind.Other;
                }

                string rawValue = kind == ExcelCellValueKind.Formula
                    ? cell.CellFormula?.Text ?? string.Empty
                    : cell.CellValue?.InnerText ?? text;
                DateTime? dateTimeValue = null;
                if (kind == ExcelCellValueKind.Number
                    && double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double serial)
                    && GetCellStyle(row, column).IsDateLike) {
                    try {
                        dateTimeValue = ExcelDateSystemConverter.FromSerial(serial, _excelDocument.DateSystem);
                        kind = ExcelCellValueKind.DateTime;
                    } catch (ArgumentException) {
                        // Retain the numeric kind when the serial cannot be represented by DateTime.
                    }
                }

                snapshot = new ExcelCellValueSnapshot(kind, text, rawValue, dataType, dateTimeValue);
                return true;
            } catch {
                return false;
            }
        }

        private void ApplyWrapText(int row, int column) {
            var cell = GetCell(row, column);
            ApplyWrapText(cell);
        }

        private void ApplyAutomaticCellFormatting(Cell cell, object? value, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>? dataType) {
            if (!RequiresAutomaticCellFormatting(value, dataType)) {
                return;
            }

            bool wroteNumber = dataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;

            // Automatically apply date format for DateTime values
            // Using Excel's built-in date format code 14 (invariant short date)
            if (wroteNumber && (value is DateTime || value is DateTimeOffset)) {
                ApplyBuiltInNumberFormat(cell, 14);
            }

            if (value is TimeSpan) {
                // Built-in format 46 renders durations using the invariant [h]:mm:ss pattern
                ApplyBuiltInNumberFormat(cell, 46);
            }

            // Enable wrap text when value contains new lines so Excel renders multiple lines correctly
            if (value is string s && (s.Contains("\n") || s.Contains("\r"))) {
                ApplyWrapText(cell);
            }
        }

        private static bool RequiresAutomaticCellFormatting(object? value, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>? dataType) {
            bool wroteNumber = dataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            return (wroteNumber && (value is DateTime || value is DateTimeOffset))
                || value is TimeSpan
                || value is string s && (s.Contains("\n") || s.Contains("\r"));
        }

        private void ApplyAutomaticCellFormattingForAppendedCell(
            Cell cell,
            object? value,
            EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>? dataType,
            uint baseStyleIndex,
            ref Dictionary<uint, uint>? dateStyleIndexes,
            ref Dictionary<uint, uint>? durationStyleIndexes,
            ref Dictionary<uint, uint>? wrapStyleIndexes) {
            bool wroteNumber = dataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;

            if (wroteNumber && (value is DateTime || value is DateTimeOffset)) {
                cell.StyleIndex = GetOrAddBuiltInNumberFormatStyleIndex(ref dateStyleIndexes, baseStyleIndex, 14);
                return;
            }

            if (value is TimeSpan) {
                cell.StyleIndex = GetOrAddBuiltInNumberFormatStyleIndex(ref durationStyleIndexes, baseStyleIndex, 46);
                return;
            }

            if (value is string s && (s.Contains("\n") || s.Contains("\r"))) {
                cell.StyleIndex = GetOrAddWrapTextStyleIndex(ref wrapStyleIndexes, baseStyleIndex);
            }
        }

        private uint GetOrAddBuiltInNumberFormatStyleIndex(ref Dictionary<uint, uint>? styleIndexes, uint baseStyleIndex, uint builtInFormatId) {
            styleIndexes ??= new Dictionary<uint, uint>();
            if (!styleIndexes.TryGetValue(baseStyleIndex, out uint styleIndex)) {
                styleIndex = GetOrCreateBuiltInNumberFormatStyleIndex(baseStyleIndex, builtInFormatId);
                styleIndexes[baseStyleIndex] = styleIndex;
            }

            return styleIndex;
        }

        private uint GetOrAddWrapTextStyleIndex(ref Dictionary<uint, uint>? styleIndexes, uint baseStyleIndex) {
            styleIndexes ??= new Dictionary<uint, uint>();
            if (!styleIndexes.TryGetValue(baseStyleIndex, out uint styleIndex)) {
                styleIndex = GetOrCreateWrapTextStyleIndex(baseStyleIndex);
                styleIndexes[baseStyleIndex] = styleIndex;
            }

            return styleIndex;
        }

        private uint GetOrCreateBuiltInNumberFormatStyleIndex(uint baseStyleIndex, uint builtInFormatId) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            WorkbookStylesPart? stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            var newFormat = GetBaseCellFormat(stylesheet, baseStyleIndex);
            newFormat.NumberFormatId = builtInFormatId;
            newFormat.ApplyNumberFormat = true;
            uint index = AppendOrReuseCellFormat(stylesheet, newFormat);
            stylesPart.Stylesheet.Save();
            return index;
        }

        private uint GetOrCreateWrapTextStyleIndex(uint baseStyleIndex) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            WorkbookStylesPart? stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            var newFormat = GetBaseCellFormat(stylesheet, baseStyleIndex);
            var alignment = newFormat.Alignment != null
                ? (Alignment)newFormat.Alignment.CloneNode(true)
                : new Alignment();
            alignment.WrapText = true;
            newFormat.Alignment = alignment;
            newFormat.ApplyAlignment = true;
            uint index = AppendOrReuseCellFormat(stylesheet, newFormat);
            stylesPart.Stylesheet.Save();
            return index;
        }

        private void ApplyWrapText(Cell cell) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            WorkbookStylesPart? stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            // Base on existing cell's style if present
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

            // Try to find an existing format with same base ids and WrapText enabled
            int wrapIndex = -1;
            for (int i = 0; i < cellFormats.Count; i++) {
                var cf = cellFormats[i];
                var align = cf.Alignment;
                bool wrap = align != null && align.WrapText != null && align.WrapText.Value;
                if (wrap && cf.NumberFormatId?.Value == baseFormat.NumberFormatId?.Value
                        && cf.FontId?.Value == baseFormat.FontId?.Value
                        && cf.FillId?.Value == baseFormat.FillId?.Value
                        && cf.BorderId?.Value == baseFormat.BorderId?.Value) {
                    wrapIndex = i;
                    break;
                }
            }

            if (wrapIndex == -1) {
                var newFormat = new CellFormat {
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
        /// Enables WrapText for every cell in a column within a given row range.
        /// </summary>
        /// <param name="fromRow">The first 1-based row index in the range.</param>
        /// <param name="toRow">The last 1-based row index in the range.</param>
        /// <param name="column">The 1-based column index whose cells should wrap.</param>
        public void WrapCells(int fromRow, int toRow, int column) {
            if (fromRow < 1 || toRow < fromRow || column < 1) return;
            WriteLockConditional(() => {
                for (int r = fromRow; r <= toRow; r++) {
                    ApplyWrapText(r, column);
                }
            });
        }

        /// <summary>
        /// Enables WrapText for the specified column and pins the target column width (in Excel character units).
        /// Useful when mixed with auto-fit operations so wrapped columns keep a predictable width.
        /// </summary>
        /// <param name="fromRow">The first 1-based row index in the range.</param>
        /// <param name="toRow">The last 1-based row index in the range.</param>
        /// <param name="column">The 1-based column index whose cells should wrap.</param>
        /// <param name="targetColumnWidth">The column width, in Excel character units, to enforce when wrapping.</param>
        public void WrapCells(int fromRow, int toRow, int column, double targetColumnWidth) {
            WrapCells(fromRow, toRow, column);
            if (targetColumnWidth > 0) {
                try { SetColumnWidth(column, targetColumnWidth); } catch { }
            }
        }

        /// <summary>
        /// Applies a horizontal alignment to a single cell.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to align.</param>
        /// <param name="column">The 1-based column index of the cell to align.</param>
        /// <param name="alignment">The horizontal alignment value to apply.</param>
        public void CellAlign(int row, int column, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues alignment) {
            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                ApplyCellFormatOverride(stylesheet, cell, format => {
                    var existingAlignment = format.Alignment != null
                        ? (Alignment)format.Alignment.CloneNode(true)
                        : new Alignment();
                    existingAlignment.Horizontal = alignment;
                    format.Alignment = existingAlignment;
                    format.ApplyAlignment = true;
                });

                stylesPart.Stylesheet.Save();
            });
        }

        /// <summary>
        /// Sets or clears shrink-to-fit text alignment on a single cell.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to align.</param>
        /// <param name="column">The 1-based column index of the cell to align.</param>
        /// <param name="shrinkToFit">Whether cell text should shrink horizontally to fit.</param>
        public void CellShrinkToFit(int row, int column, bool shrinkToFit = true) {
            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                ApplyCellFormatOverride(stylesheet, cell, format => {
                    var existingAlignment = format.Alignment != null
                        ? (Alignment)format.Alignment.CloneNode(true)
                        : new Alignment();
                    existingAlignment.ShrinkToFit = shrinkToFit;
                    format.Alignment = existingAlignment;
                    format.ApplyAlignment = true;
                });

                stylesPart.Stylesheet.Save();
            });
        }

        /// <summary>
        /// Applies Excel text rotation to a single cell.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to align.</param>
        /// <param name="column">The 1-based column index of the cell to align.</param>
        /// <param name="rotation">Open XML text rotation value. Use 0-90 for upward rotation, 91-180 for downward rotation, or 255 for stacked vertical text.</param>
        public void CellTextRotation(int row, int column, int rotation) {
            if ((rotation < 0 || rotation > 180) && rotation != 255) {
                throw new ArgumentOutOfRangeException(nameof(rotation), "Text rotation must be between 0 and 180, or 255 for stacked vertical text.");
            }

            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                ApplyCellFormatOverride(stylesheet, cell, format => {
                    var existingAlignment = format.Alignment != null
                        ? (Alignment)format.Alignment.CloneNode(true)
                        : new Alignment();
                    existingAlignment.TextRotation = (UInt32Value)(uint)rotation;
                    format.Alignment = existingAlignment;
                    format.ApplyAlignment = true;
                });

                stylesPart.Stylesheet.Save();
            });
        }

        /// <summary>
        /// Applies a vertical alignment to a single cell.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to align.</param>
        /// <param name="column">The 1-based column index of the cell to align.</param>
        /// <param name="alignment">The vertical alignment value to apply.</param>
        public void CellVerticalAlign(int row, int column, VerticalAlignmentValues alignment) {
            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                ApplyCellFormatOverride(stylesheet, cell, format => {
                    var existingAlignment = format.Alignment != null
                        ? (Alignment)format.Alignment.CloneNode(true)
                        : new Alignment();
                    existingAlignment.Vertical = alignment;
                    format.Alignment = existingAlignment;
                    format.ApplyAlignment = true;
                });

                stylesPart.Stylesheet.Save();
            });
        }

        /// <summary>
        /// Applies the same border style to all sides of a single cell.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to style.</param>
        /// <param name="column">The 1-based column index of the cell to style.</param>
        /// <param name="style">The border style to apply on all four sides.</param>
        /// <param name="hexColor">Optional border color expressed as ARGB or RGB hex.</param>
        public void CellBorder(int row, int column, BorderStyleValues style, string? hexColor = null) {
            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                var baseFormat = GetBaseCellFormat(stylesheet, cell.StyleIndex?.Value ?? 0U);
                var borderId = GetOrCreateBorderVariant(stylesheet, GetOptionalValue(baseFormat.BorderId), border => SetUniformBorder(border, style, hexColor));
                ApplyCellFormatOverride(stylesheet, cell, format => {
                    format.BorderId = borderId;
                    format.ApplyBorder = true;
                });

                stylesPart.Stylesheet.Save();
            });
        }

        /// <summary>
        /// Applies diagonal border lines to a single cell.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to style.</param>
        /// <param name="column">The 1-based column index of the cell to style.</param>
        /// <param name="style">The diagonal border style.</param>
        /// <param name="hexColor">Optional border color expressed as ARGB or RGB hex.</param>
        /// <param name="diagonalUp">Whether to draw the bottom-left to top-right diagonal.</param>
        /// <param name="diagonalDown">Whether to draw the top-left to bottom-right diagonal.</param>
        public void CellDiagonalBorder(int row, int column, BorderStyleValues style, string? hexColor = null, bool diagonalUp = true, bool diagonalDown = true) {
            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                var baseFormat = GetBaseCellFormat(stylesheet, cell.StyleIndex?.Value ?? 0U);
                var borderId = GetOrCreateBorderVariant(stylesheet, GetOptionalValue(baseFormat.BorderId), border => SetDiagonalBorder(border, style, hexColor, diagonalUp, diagonalDown));
                ApplyCellFormatOverride(stylesheet, cell, format => {
                    format.BorderId = borderId;
                    format.ApplyBorder = true;
                });

                stylesPart.Stylesheet.Save();
            });
        }

        /// <summary>
        /// Applies a font color (ARGB hex or #RRGGBB) to a single cell.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to recolor.</param>
        /// <param name="column">The 1-based column index of the cell to recolor.</param>
        /// <param name="hexColor">The desired font color expressed as an ARGB or RGB hex string.</param>
        public void CellFontColor(int row, int column, string hexColor) {
            if (string.IsNullOrWhiteSpace(hexColor)) return;
            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                string argb = NormalizeHexColor(hexColor);

                uint baseIndex = cell.StyleIndex?.Value ?? 0U;
                var baseFormat = GetBaseCellFormat(stylesheet, baseIndex);
                var fontId = GetOrCreateFontVariant(stylesheet, GetOptionalValue(baseFormat.FontId), font => SetFontColor(font, argb));
                ApplyCellFormatOverride(stylesheet, cell, format => {
                    format.FontId = fontId;
                    format.ApplyFont = true;
                });

                stylesPart.Stylesheet.Save();
            });
        }

        private void ApplyFontBold(Cell cell, bool bold) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null)
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            uint baseIndex = cell.StyleIndex?.Value ?? 0U;
            var baseFormat = GetBaseCellFormat(stylesheet, baseIndex);
            var boldFontId = GetOrCreateFontVariant(stylesheet, GetOptionalValue(baseFormat.FontId), font => SetBold(font, bold));
            ApplyCellFormatOverride(stylesheet, cell, format => {
                format.FontId = boldFontId;
                format.ApplyFont = true;
            });
            stylesPart.Stylesheet.Save();
        }

        private void ApplyFontItalic(Cell cell, bool italic) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null)
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            uint baseIndex = cell.StyleIndex?.Value ?? 0U;
            var baseFormat = GetBaseCellFormat(stylesheet, baseIndex);
            var italicFontId = GetOrCreateFontVariant(stylesheet, GetOptionalValue(baseFormat.FontId), font => SetItalic(font, italic));
            ApplyCellFormatOverride(stylesheet, cell, format => {
                format.FontId = italicFontId;
                format.ApplyFont = true;
            });
            stylesPart.Stylesheet.Save();
        }

        private void ApplyFontUnderline(Cell cell, bool underline) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null)
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            uint baseIndex = cell.StyleIndex?.Value ?? 0U;
            var baseFormat = GetBaseCellFormat(stylesheet, baseIndex);
            var underlineFontId = GetOrCreateFontVariant(stylesheet, GetOptionalValue(baseFormat.FontId), font => SetUnderline(font, underline));
            ApplyCellFormatOverride(stylesheet, cell, format => {
                format.FontId = underlineFontId;
                format.ApplyFont = true;
            });
            stylesPart.Stylesheet.Save();
        }

        private void ApplyFontName(Cell cell, string fontName) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null)
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            uint baseIndex = cell.StyleIndex?.Value ?? 0U;
            var baseFormat = GetBaseCellFormat(stylesheet, baseIndex);
            var namedFontId = GetOrCreateFontVariant(stylesheet, GetOptionalValue(baseFormat.FontId), font => SetFontName(font, fontName));
            ApplyCellFormatOverride(stylesheet, cell, format => {
                format.FontId = namedFontId;
                format.ApplyFont = true;
            });
            stylesPart.Stylesheet.Save();
        }

        private void ApplyFontSize(Cell cell, double fontSize) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null)
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            uint baseIndex = cell.StyleIndex?.Value ?? 0U;
            var baseFormat = GetBaseCellFormat(stylesheet, baseIndex);
            var fontId = GetOrCreateFontVariant(stylesheet, GetOptionalValue(baseFormat.FontId), font => SetFontSize(font, fontSize));
            ApplyCellFormatOverride(stylesheet, cell, format => {
                format.FontId = fontId;
                format.ApplyFont = true;
            });
            stylesPart.Stylesheet.Save();
        }

        private static string NormalizeHexColor(string hex) {
            hex = hex.Trim();
            if (hex.StartsWith("#")) hex = hex.Substring(1);
            if (hex.Length == 6) return "FF" + hex.ToUpperInvariant();
            if (hex.Length == 8) return hex.ToUpperInvariant();
            // Fallback default
            return "FFFFFFFF";
        }

        private void ApplyBackground(Cell cell, string hexColor) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null)
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            // Create a fill with solid color
            string argb = NormalizeHexColor(hexColor);
            var fill = new Fill(new PatternFill {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = argb },
                BackgroundColor = new BackgroundColor { Rgb = argb }
            });
            var fillId = GetOrCreateFill(stylesheet, fill);
            ApplyCellFormatOverride(stylesheet, cell, format => {
                format.FillId = fillId;
                format.ApplyFill = true;
            });
            stylesPart.Stylesheet.Save();
        }

        private void FillRangeCore(int firstRow, int firstColumn, int lastRow, int lastColumn, string hexColor) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null)
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            string argb = NormalizeHexColor(hexColor);
            var fill = new Fill(new PatternFill {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = argb },
                BackgroundColor = new BackgroundColor { Rgb = argb }
            });
            uint fillId = GetOrCreateFill(stylesheet, fill);
            var styleIndexes = new Dictionary<uint, uint>();

            for (int row = firstRow; row <= lastRow; row++) {
                for (int column = firstColumn; column <= lastColumn; column++) {
                    Cell cell = GetCell(row, column);
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    cell.StyleIndex = GetOrAddCellFormatOverride(styleIndexes, stylesheet, baseStyleIndex, format => {
                        format.FillId = fillId;
                        format.ApplyFill = true;
                    });
                }
            }

            stylesPart.Stylesheet.Save();
        }

        private void ApplyBuiltInNumberFormat(int row, int column, uint builtInFormatId) {
            Cell cell = GetCell(row, column);
            ApplyBuiltInNumberFormat(cell, builtInFormatId);
        }

        private void ApplyBuiltInNumberFormat(Cell cell, uint builtInFormatId) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            WorkbookStylesPart? stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            ApplyCellFormatOverride(stylesheet, cell, format => {
                format.NumberFormatId = builtInFormatId;
                format.ApplyNumberFormat = true;
            });
            stylesPart.Stylesheet.Save();
        }

        private void FormatCellCore(int row, int column, string numberFormat) {
            Cell cell = GetCell(row, column);

            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            WorkbookStylesPart? stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            uint numberFormatId = GetOrCreateNumberFormatId(stylesheet, numberFormat);

            ApplyCellFormatOverride(stylesheet, cell, format => {
                format.NumberFormatId = numberFormatId;
                format.ApplyNumberFormat = true;
            });
            stylesPart.Stylesheet.Save();
        }

        private void FormatRangeCore(int firstRow, int firstColumn, int lastRow, int lastColumn, string numberFormat) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            WorkbookStylesPart? stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            uint numberFormatId = GetOrCreateNumberFormatId(stylesheet, numberFormat);
            var styleIndexes = new Dictionary<uint, uint>();

            for (int row = firstRow; row <= lastRow; row++) {
                for (int column = firstColumn; column <= lastColumn; column++) {
                    Cell cell = GetCell(row, column);
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    cell.StyleIndex = GetOrAddCellFormatOverride(styleIndexes, stylesheet, baseStyleIndex, format => {
                        format.NumberFormatId = numberFormatId;
                        format.ApplyNumberFormat = true;
                    });
                }
            }

            stylesPart.Stylesheet.Save();
        }

    }
}
