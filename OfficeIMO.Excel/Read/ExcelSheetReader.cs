using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Read
{
    /// <summary>
    /// Reader for a single worksheet. Offers enumeration and conversion helpers.
    /// </summary>
    public sealed partial class ExcelSheetReader
    {
        private readonly string _sheetName;
        private readonly WorksheetPart _wsPart;
        private readonly SharedStringCache _sst;
        private readonly StylesCache _styles;
        private readonly ExcelReadOptions _opt;

        internal ExcelSheetReader(string sheetName, WorksheetPart wsPart, SharedStringCache sst, StylesCache styles, ExcelReadOptions opt)
        {
            _sheetName = sheetName;
            _wsPart = wsPart;
            _sst = sst;
            _styles = styles;
            _opt = opt;
        }

        /// <summary>
        /// Worksheet name.
        /// </summary>
        public string Name => _sheetName;

        /// <summary>
        /// Enumerates non-empty cells as (Row, Column, Value). Values are typed when possible.
        /// </summary>
        public IEnumerable<CellValueInfo> EnumerateCells()
        {
            var sheetData = _wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) yield break;

            foreach (var row in sheetData.Elements<Row>())
            {
                var rIndex = checked((int)row.RowIndex!.Value);
                foreach (var cell in row.Elements<Cell>())
                {
                    var (cIndex, _) = A1.ParseCellRef(cell.CellReference?.Value ?? string.Empty);
                    var value = ConvertCell(cell);
                    if (value is not null || CellHasExplicitBlank(cell))
                        yield return new CellValueInfo(rIndex, cIndex, value);
                }
            }
        }

        // ---------- Internals ----------

        private static bool CellHasExplicitBlank(Cell cell)
        {
            return cell.CellValue is not null && string.IsNullOrEmpty(cell.CellValue.InnerText);
        }

        private static string? ExtractRawText(Cell cell)
        {
            if (cell.CellFormula is not null && cell.CellValue is not null) return cell.CellValue.InnerText;
            if (cell.CellValue is not null) return cell.CellValue.InnerText;
            return null;
        }

        private static string? ExtractInlineString(Cell cell)
        {
            var inline = cell.InlineString;
            if (inline?.Text?.Text != null) return inline.Text.Text;
            if (inline?.HasChildren == true)
            {
                var runs = inline.Elements<Run>().Select(r => r.Text?.Text ?? string.Empty);
                return string.Concat(runs);
            }
            return null;
        }

        private object? ConvertCell(Cell cell)
        {
            var raw = new CellRaw
            {
                Row = (int)(cell.CellReference != null ? A1.ParseCellRef(cell.CellReference.Value).Row : 0),
                Col = (int)(cell.CellReference != null ? A1.ParseCellRef(cell.CellReference.Value).Col : 0),
                TypeHint = cell.DataType?.Value,
                StyleIndex = cell.StyleIndex?.Value,
                HasFormula = cell.CellFormula is not null,
                RawText = ExtractRawText(cell),
                InlineText = ExtractInlineString(cell)
            };
            return ConvertRaw(raw).TypedValue;
        }

        private CellRaw ConvertRaw(CellRaw raw)
        {
            if (raw.HasFormula)
            {
                if (_opt.UseCachedFormulaResult && raw.RawText != null)
                {
                    raw.TypedValue = ConvertByHints(raw.TypeHint, raw.StyleIndex, raw.RawText, raw.InlineText);
                }
                else
                {
                    raw.TypedValue = raw.RawText ?? raw.InlineText; // return formula/cached text
                }
                return raw;
            }

            raw.TypedValue = ConvertByHints(raw.TypeHint, raw.StyleIndex, raw.RawText, raw.InlineText);
            return raw;
        }

        private object? ConvertByHints(EnumValue<CellValues>? type, uint? styleIndex, string? rawText, string? inlineText)
        {
            // Custom converter hook (cell-level). If provided and handled, honor it.
            var hook = _opt.CellValueConverter;
            if (hook != null)
            {
                var ctx = new ExcelCellContext(type?.Value, styleIndex, rawText, inlineText, _opt.Culture);
                var res = hook(ctx);
                if (res.Handled) return res.Value;
            }
            if (!string.IsNullOrEmpty(inlineText)) return inlineText;

            if (type?.Value == CellValues.SharedString && int.TryParse(rawText, System.Globalization.NumberStyles.Integer, CultureInfo.InvariantCulture, out var sstIndex))
                return _sst.Get(sstIndex);

            if (type?.Value == CellValues.Boolean && rawText != null)
                return rawText == "1";

            if (type?.Value == CellValues.Number && rawText != null)
            {
                if (_opt.TreatDatesUsingNumberFormat && styleIndex is not null && _styles.IsDateLike(styleIndex.Value))
                {
                    if (double.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var oa))
                        return System.DateTime.FromOADate(oa);
                }
                if (double.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, _opt.Culture, out var num))
                    return num;
                return rawText;
            }

            if (type?.Value == CellValues.Date && rawText != null)
            {
                if (System.DateTime.TryParse(rawText, _opt.Culture, System.Globalization.DateTimeStyles.AssumeLocal, out var dt))
                    return dt;
                return rawText;
            }

            if (type?.Value == CellValues.String || type?.Value == CellValues.InlineString || type?.Value == CellValues.SharedString)
            {
                if (type?.Value == CellValues.SharedString && rawText != null && int.TryParse(rawText, out var idx))
                    return _sst.Get(idx);
                return rawText ?? inlineText;
            }

            if (rawText != null)
            {
                if (_opt.TreatDatesUsingNumberFormat && styleIndex is not null && _styles.IsDateLike(styleIndex.Value))
                {
                    if (double.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var oa))
                        return System.DateTime.FromOADate(oa);
                    return rawText;
                }

                if (double.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, _opt.Culture, out var num))
                    return num;
                return rawText;
            }

            return null;
        }

        private struct CellRaw
        {
            public int Row;
            public int Col;
            public EnumValue<CellValues>? TypeHint;
            public uint? StyleIndex;
            public bool HasFormula;
            public string? RawText;
            public string? InlineText;
            public object? TypedValue;
        }
    }
}
