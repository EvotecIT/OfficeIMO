using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private Dictionary<uint, uint>? _cellValueDateStyleIndexes;
        private Dictionary<uint, uint>? _cellValueDurationStyleIndexes;
        private uint? _cellValueDefaultDateStyleIndex;
        private uint? _cellValueDefaultDurationStyleIndex;

        // Core implementation: single source of truth (no locks here)
        private void CellValueCore(int row, int column, object? value) {
            switch (value) {
                case string text:
                    CellStringValueCore(row, column, text);
                    return;
                case double number:
                    CellDoubleValueCore(row, column, number);
                    return;
                case float number:
                    CellDoubleValueCore(row, column, Convert.ToDouble(number));
                    return;
                case decimal number:
                    CellDecimalValueCore(row, column, number);
                    return;
                case int number:
                    CellNumberTextValueCore(row, column, number.ToString(CultureInfo.InvariantCulture));
                    return;
                case long number:
                    CellNumberTextValueCore(row, column, number.ToString(CultureInfo.InvariantCulture));
                    return;
                case short number:
                    CellNumberTextValueCore(row, column, number.ToString(CultureInfo.InvariantCulture));
                    return;
                case uint number:
                    CellNumberTextValueCore(row, column, number.ToString(CultureInfo.InvariantCulture));
                    return;
                case ulong number:
                    CellNumberTextValueCore(row, column, number.ToString(CultureInfo.InvariantCulture));
                    return;
                case ushort number:
                    CellNumberTextValueCore(row, column, number.ToString(CultureInfo.InvariantCulture));
                    return;
                case byte number:
                    CellNumberTextValueCore(row, column, number.ToString(CultureInfo.InvariantCulture));
                    return;
                case sbyte number:
                    CellNumberTextValueCore(row, column, number.ToString(CultureInfo.InvariantCulture));
                    return;
                case bool boolean:
                    CellBooleanValueCore(row, column, boolean);
                    return;
                case DateTime dateTime:
                    CellDateTimeValueCore(row, column, dateTime);
                    return;
                case DateTimeOffset dateTimeOffset:
                    CellDateTimeOffsetValueCore(row, column, dateTimeOffset);
                    return;
                case TimeSpan timeSpan:
                    CellTimeSpanValueCore(row, column, timeSpan);
                    return;
            }

            var (cellValue, dataType) = CoerceForCell(value);

            var cell = GetCell(row, column);
            cell.CellValue = cellValue;
            cell.DataType = dataType;
            ApplyAutomaticCellFormatting(cell, value, dataType);
            ClearHeaderCache();
        }

        private void CellStringValueCore(int row, int column, string? value) {
            if (value == null) {
                CellValueCore(row, column, value);
                return;
            }

            CoerceValueHelper.ValidateSharedStringLength(value, nameof(value));
            int sharedStringIndex = _excelDocument.GetSharedStringIndex(value);

            var cell = GetCell(row, column);
            SetExistingCellSharedStringValue(cell, value, sharedStringIndex);
            ClearHeaderCache();
        }

        private void SetExistingCellSharedStringValue(Cell cell, string value, int sharedStringIndex) {
            cell.CellValue = new CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture));
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString;
            if (value.Contains("\n") || value.Contains("\r")) {
                ApplyWrapText(cell);
            }
        }

        private void CellDoubleValueCore(int row, int column, double value) {
            var cell = GetCell(row, column);
            cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            ClearHeaderCache();
        }

        private void CellDecimalValueCore(int row, int column, decimal value) {
            var cell = GetCell(row, column);
            cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            ClearHeaderCache();
        }

        private void CellNumberTextValueCore(int row, int column, string text) {
            var cell = GetCell(row, column);
            cell.CellValue = new CellValue(text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            ClearHeaderCache();
        }

        private void CellBooleanValueCore(int row, int column, bool value) {
            var cell = GetCell(row, column);
            cell.CellValue = new CellValue(value ? "1" : "0");
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean;
            ClearHeaderCache();
        }

        private void CellDateTimeValueCore(int row, int column, DateTime value) {
            double serial = value.ToOADate();
            var cell = GetCell(row, column);
            uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
            cell.CellValue = new CellValue(serial.ToString(CultureInfo.InvariantCulture));
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            cell.StyleIndex = baseStyleIndex == 0U
                ? (_cellValueDefaultDateStyleIndex ??= GetOrCreateBuiltInNumberFormatStyleIndex(0U, 14))
                : GetOrAddBuiltInNumberFormatStyleIndex(ref _cellValueDateStyleIndexes, baseStyleIndex, 14);
            ClearHeaderCache();
        }

        private void CellDateTimeOffsetValueCore(int row, int column, DateTimeOffset value) {
            var dateTimeOffsetStrategy = _excelDocument.DateTimeOffsetWriteStrategy;
            var (cellValue, cellType) = CoerceValueHelper.HandleDateTimeOffset(
                value,
                s => {
                    int idx = _excelDocument.GetSharedStringIndex(s);
                    return new CellValue(idx.ToString(CultureInfo.InvariantCulture));
                },
                dateTimeOffsetStrategy,
                nameof(value));

            var cell = GetCell(row, column);
            cell.CellValue = cellValue;
            cell.DataType = cellType;
            if (cellType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number) {
                uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                cell.StyleIndex = baseStyleIndex == 0U
                    ? (_cellValueDefaultDateStyleIndex ??= GetOrCreateBuiltInNumberFormatStyleIndex(0U, 14))
                    : GetOrAddBuiltInNumberFormatStyleIndex(ref _cellValueDateStyleIndexes, baseStyleIndex, 14);
            }

            ClearHeaderCache();
        }

        private void CellFormulaCore(int row, int column, string formula) {
            Cell cell = GetCell(row, column);
            // Excel formulas in XML should not start with '=' and must not include illegal control characters
            var safe = Utilities.ExcelSanitizer.SanitizeFormula(formula);
            cell.CellFormula = new CellFormula(safe);
            ClearHeaderCache();
        }

        private void CellTimeSpanValueCore(int row, int column, TimeSpan value) {
            double serial = value.TotalDays;
            var cell = GetCell(row, column);
            uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
            cell.CellValue = new CellValue(serial.ToString(CultureInfo.InvariantCulture));
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            cell.StyleIndex = baseStyleIndex == 0U
                ? (_cellValueDefaultDurationStyleIndex ??= GetOrCreateBuiltInNumberFormatStyleIndex(0U, 46))
                : GetOrAddBuiltInNumberFormatStyleIndex(ref _cellValueDurationStyleIndexes, baseStyleIndex, 46);
            ClearHeaderCache();
        }

        // Core coercion logic shared between sequential and parallel operations
        private (CellValue cellValue, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> dataType) CoerceForCell(object? value) {
            var dateTimeOffsetStrategy = _excelDocument.DateTimeOffsetWriteStrategy;
            var (cellValue, cellType) = CoerceValueHelper.Coerce(
                value,
                s => {
                    int idx = _excelDocument.GetSharedStringIndex(s);
                    return new CellValue(idx.ToString(CultureInfo.InvariantCulture));
                },
                dateTimeOffsetStrategy);
            return (cellValue, new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(cellType));
        }



        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, string value) {
            if (_isBatchOperation || Locking.IsNoLock) {
                CellStringValueCore(row, column, value);
                MarkRequiresSavePreparation();
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellStringValueCore(row, column, value);
                MarkRequiresSavePreparation();
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, double value) {
            if (_isBatchOperation || Locking.IsNoLock) {
                CellDoubleValueCore(row, column, value);
                MarkRequiresSavePreparation();
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellDoubleValueCore(row, column, value);
                MarkRequiresSavePreparation();
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, decimal value) {
            if (_isBatchOperation || Locking.IsNoLock) {
                CellDecimalValueCore(row, column, value);
                MarkRequiresSavePreparation();
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellDecimalValueCore(row, column, value);
                MarkRequiresSavePreparation();
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, DateTime value) {
            if (_isBatchOperation || Locking.IsNoLock) {
                CellDateTimeValueCore(row, column, value);
                MarkRequiresSavePreparation();
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellDateTimeValueCore(row, column, value);
                MarkRequiresSavePreparation();
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, DateTimeOffset value) {
            WriteLockConditional(() => CellValueCore(row, column, value));
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, TimeSpan value) {
            if (_isBatchOperation || Locking.IsNoLock) {
                CellTimeSpanValueCore(row, column, value);
                MarkRequiresSavePreparation();
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellTimeSpanValueCore(row, column, value);
                MarkRequiresSavePreparation();
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, uint value) {
            WriteLockConditional(() => CellValueCore(row, column, value));
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, ulong value) {
            WriteLockConditional(() => CellValueCore(row, column, value));
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, ushort value) {
            WriteLockConditional(() => CellValueCore(row, column, value));
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, byte value) {
            WriteLockConditional(() => CellValueCore(row, column, value));
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, sbyte value) {
            WriteLockConditional(() => CellValueCore(row, column, value));
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, bool value) {
            if (_isBatchOperation || Locking.IsNoLock) {
                CellBooleanValueCore(row, column, value);
                MarkRequiresSavePreparation();
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellBooleanValueCore(row, column, value);
                MarkRequiresSavePreparation();
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <summary>
        /// Sets a formula in the specified cell.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="formula">The formula expression.</param>
        public void CellFormula(int row, int column, string formula) {
            if (_isBatchOperation || Locking.IsNoLock) {
                CellFormulaCore(row, column, formula);
                MarkRequiresSavePreparation();
                return;
            }

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                _excelDocument.MaterializeDeferredDataSetImport();
            }

            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellFormulaCore(row, column, formula);
                MarkRequiresSavePreparation();
            } finally {
                lck.ExitWriteLock();
            }
        }

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
                _excelDocument.MaterializeDeferredDataSetImport();
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
                    _excelDocument.MaterializeDeferredDataSetImport();
                }

                var cell = GetCell(row, column);
                if (cell == null) return false;
                // Resolve shared string if needed
                if (cell.DataType != null && cell.DataType.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                    if (TryParseCellTextSharedStringIndex(cell.InnerText, out int ssid)) {
                        string? sharedText = GetCellTextSharedStringCache().Get(ssid);
                        if (sharedText != null) {
                            text = sharedText;
                            return true;
                        }

                        return false;
                    }
                }
                // Otherwise, return inner text (numbers/booleans as invariant string)
                text = cell.InnerText ?? string.Empty;
                return !string.IsNullOrEmpty(text);
            } catch { return false; }
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

            ApplyCellFormatOverride(stylesheet, cell, format => {
                format.NumberFormatId = numberFormatId;
                format.ApplyNumberFormat = true;
            });
            stylesPart.Stylesheet.Save();
        }

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

        private static void SetUniformBorder(Border border, BorderStyleValues style, string? hexColor) {
            var argb = string.IsNullOrWhiteSpace(hexColor) ? null : NormalizeHexColor(hexColor!);
            border.LeftBorder = CreateBorderSide<LeftBorder>(style, argb);
            border.RightBorder = CreateBorderSide<RightBorder>(style, argb);
            border.TopBorder = CreateBorderSide<TopBorder>(style, argb);
            border.BottomBorder = CreateBorderSide<BottomBorder>(style, argb);
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

        /// <summary>
        /// Sets the specified value into a cell, inferring the data type.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="value">The value to assign.</param>
        public void CellValue(int row, int column, object? value) {
            if (_isBatchOperation || Locking.IsNoLock) {
                CellValueCore(row, column, value);
                MarkRequiresSavePreparation();
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellValueCore(row, column, value);
                MarkRequiresSavePreparation();
            } finally {
                lck.ExitWriteLock();
            }
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
