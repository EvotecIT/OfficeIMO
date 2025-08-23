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
        }

        // Compute-only coercion (no OpenXML mutations, except SharedString for now)
        private (CellValue cellValue, EnumValue<CellValues> dataType) CoerceForCell(object value)
        {
            switch (value)
            {
                case null:
                    return (new CellValue(string.Empty), new EnumValue<CellValues>(CellValues.String));
                case string s:
                    // TODO: SharedString index should be resolved via planner in parallel scenarios
                    int sharedStringIndex = _excelDocument.GetSharedStringIndex(s);
                    return (new CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.SharedString));
                case double d:
                    return (new CellValue(d.ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                case float f:
                    return (new CellValue(Convert.ToDouble(f).ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                case decimal dec:
                    return (new CellValue(dec.ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                case int i:
                    return (new CellValue(((double)i).ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                case long l:
                    return (new CellValue(((double)l).ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                case DateTime dt:
                    return (new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                case DateTimeOffset dto:
                    return (new CellValue(dto.UtcDateTime.ToOADate().ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                case TimeSpan ts:
                    return (new CellValue(ts.TotalDays.ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                case bool b:
                    return (new CellValue(b ? "1" : "0"), new EnumValue<CellValues>(CellValues.Boolean));
                case uint ui:
                    return (new CellValue(((double)ui).ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                case ulong ul:
                    return (new CellValue(((double)ul).ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                case ushort us:
                    return (new CellValue(((double)us).ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                case byte by:
                    return (new CellValue(((double)by).ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                case sbyte sb:
                    return (new CellValue(((double)sb).ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                case short sh:
                    return (new CellValue(((double)sh).ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.Number));
                default:
                    string stringValue = value?.ToString() ?? string.Empty;
                    int defaultIndex = _excelDocument.GetSharedStringIndex(stringValue);
                    return (new CellValue(defaultIndex.ToString(CultureInfo.InvariantCulture)), new EnumValue<CellValues>(CellValues.SharedString));
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
                CellFormula(row, column, formula);
            }
            if (!string.IsNullOrEmpty(numberFormat)) {
                FormatCell(row, column, numberFormat);
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

        private void ApplyBuiltInNumberFormat(int row, int column, uint builtInFormatId) {
            Cell cell = GetCell(row, column);

            WorkbookStylesPart stylesPart = _excelDocument._spreadSheetDocument.WorkbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = _excelDocument._spreadSheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();

            stylesheet.Fonts ??= new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font());
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

            stylesheet.Fills ??= new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }));
            stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();

            stylesheet.Borders ??= new Borders(new DocumentFormat.OpenXml.Spreadsheet.Border());
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();

            stylesheet.CellStyleFormats ??= new CellStyleFormats(new CellFormat());
            stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.Count();

            stylesheet.CellFormats ??= new CellFormats(new CellFormat());

            var cellFormats = stylesheet.CellFormats.Elements<CellFormat>().ToList();
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

            WorkbookStylesPart stylesPart = _excelDocument._spreadSheetDocument.WorkbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = _excelDocument._spreadSheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();

                stylesheet.Fonts ??= new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font());
                stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

                stylesheet.Fills ??= new Fills(new Fill());
                stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();

                stylesheet.Borders ??= new Borders(new Border());
                stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();

                stylesheet.CellStyleFormats ??= new CellStyleFormats(new CellFormat());
                stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.Count();

                stylesheet.CellFormats ??= new CellFormats(new CellFormat());
                if (stylesheet.CellFormats.Count == null || stylesheet.CellFormats.Count.Value == 0) {
                    stylesheet.CellFormats.Count = 1;
                }

                stylesheet.NumberingFormats ??= new NumberingFormats();

                NumberingFormat existingFormat = stylesheet.NumberingFormats.Elements<NumberingFormat>()
                    .FirstOrDefault(n => n.FormatCode != null && n.FormatCode.Value == numberFormat);

                uint numberFormatId;
                if (existingFormat != null) {
                    numberFormatId = existingFormat.NumberFormatId.Value;
                } else {
                    numberFormatId = stylesheet.NumberingFormats.Elements<NumberingFormat>().Any()
                        ? stylesheet.NumberingFormats.Elements<NumberingFormat>().Max(n => n.NumberFormatId.Value) + 1
                        : 164U;
                    NumberingFormat numberingFormat = new NumberingFormat {
                        NumberFormatId = numberFormatId,
                        FormatCode = StringValue.FromString(numberFormat)
                    };
                    stylesheet.NumberingFormats.Append(numberingFormat);
                    stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Count();
                }

                var cellFormats = stylesheet.CellFormats.Elements<CellFormat>().ToList();
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
                    stylesheet.CellFormats.Append(cellFormat);
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    formatIndex = cellFormats.Count;
                }

                cell.StyleIndex = (uint)formatIndex;
                stylesPart.Stylesheet.Save();
        }

        /// <summary>
        /// Sets the value of a cell.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="value">The value to assign.</param>
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

