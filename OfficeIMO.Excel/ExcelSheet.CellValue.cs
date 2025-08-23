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

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, string value) {
            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                int sharedStringIndex = _excelDocument.GetSharedStringIndex(value);
                cell.CellValue = new CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.SharedString;
            });
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, double value) {
            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
            });
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, decimal value) {
            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
            });
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, DateTime value) {
            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value.ToOADate().ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
            });
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, DateTimeOffset value) {
            CellValue(row, column, value.UtcDateTime);
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, TimeSpan value) {
            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value.TotalDays.ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
            });
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
            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value ? "1" : "0");
                cell.DataType = CellValues.Boolean;
            });
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
            WriteLock(() => {
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
                        ApplyNumberFormat = true
                    };
                    stylesheet.CellFormats.Append(cellFormat);
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    formatIndex = cellFormats.Count;
                }

                cell.StyleIndex = (uint)formatIndex;
                stylesPart.Stylesheet.Save();
            });
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
            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                switch (value) {
                    case string s:
                        int sharedStringIndex = _excelDocument.GetSharedStringIndex(s);
                        cell.CellValue = new CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.SharedString;
                        break;
                    case double d:
                        cell.CellValue = new CellValue(d.ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    case float f:
                        cell.CellValue = new CellValue(Convert.ToDouble(f).ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    case decimal dec:
                        cell.CellValue = new CellValue(dec.ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    case int i:
                        cell.CellValue = new CellValue(((double)i).ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    case long l:
                        cell.CellValue = new CellValue(((double)l).ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    case DateTime dt:
                        cell.CellValue = new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    case DateTimeOffset dto:
                        cell.CellValue = new CellValue(dto.UtcDateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    case TimeSpan ts:
                        cell.CellValue = new CellValue(ts.TotalDays.ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    case bool b:
                        cell.CellValue = new CellValue(b ? "1" : "0");
                        cell.DataType = CellValues.Boolean;
                        break;
                    case uint ui:
                        cell.CellValue = new CellValue(((double)ui).ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    case ulong ul:
                        cell.CellValue = new CellValue(((double)ul).ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    case ushort us:
                        cell.CellValue = new CellValue(((double)us).ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    case byte by:
                        cell.CellValue = new CellValue(((double)by).ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    case sbyte sb:
                        cell.CellValue = new CellValue(((double)sb).ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    case short sh:
                        cell.CellValue = new CellValue(((double)sh).ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.Number;
                        break;
                    default:
                        string stringValue = value?.ToString() ?? string.Empty;
                        int defaultIndex = _excelDocument.GetSharedStringIndex(stringValue);
                        cell.CellValue = new CellValue(defaultIndex.ToString(CultureInfo.InvariantCulture));
                        cell.DataType = CellValues.SharedString;
                        break;
                }
            });
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

