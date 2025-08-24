using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel
{
    /// <summary>
    /// Data operations: autofilter, sorting (values-only), find/replace, and validation.
    /// Kept small and focused; advanced formatting lives in dedicated files.
    /// </summary>
    public partial class ExcelSheet
    {
        // -------- AutoFilter --------

        /// <summary>
        /// Attaches an AutoFilter to the given A1 range (e.g., "A1:C200").
        /// </summary>
        public void AutoFilterAdd(string a1Range)
        {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            WriteLock(() =>
            {
                var ws = _worksheetPart.Worksheet;
                // Remove existing AutoFilter, if any
                var existing = ws.Elements<AutoFilter>().FirstOrDefault();
                existing?.Remove();

                ws.InsertAfter(new AutoFilter { Reference = a1Range }, ws.GetFirstChild<SheetData>());
                ws.Save();
            });
        }

        /// <summary>
        /// Clears any AutoFilter from the worksheet.
        /// </summary>
        public void AutoFilterClear()
        {
            WriteLock(() =>
            {
                var ws = _worksheetPart.Worksheet;
                var existing = ws.Elements<AutoFilter>().FirstOrDefault();
                existing?.Remove();
                ws.Save();
            });
        }

        /// <summary>
        /// Applies an AutoFilter equals filter to a column resolved by header within the current AutoFilter range.
        /// Ensures an AutoFilter exists over the sheet's UsedRange when none is present.
        /// </summary>
        public void AutoFilterByHeaderEquals(string header, IEnumerable<string> values)
        {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            if (values == null) throw new ArgumentNullException(nameof(values));

            WriteLock(() =>
            {
                var ws = _worksheetPart.Worksheet;
                var af = ws.Elements<AutoFilter>().FirstOrDefault();
                if (af == null)
                {
                    var used = GetUsedRangeA1();
                    af = new AutoFilter { Reference = used };
                    ws.InsertAfter(af, ws.GetFirstChild<SheetData>());
                }

                var (r1, c1, r2, c2) = Read.A1.ParseRange(af.Reference!);
                int colIndex = ColumnIndexByHeader(header);
                if (colIndex < c1 || colIndex > c2)
                    throw new ArgumentOutOfRangeException(nameof(header), $"Header '{header}' is outside the AutoFilter range {af.Reference}.");

                // ColumnId is zero-based within the AutoFilter range
                uint columnId = (uint)(colIndex - c1);

                // Remove existing filter for this ColumnId
                var existingColumn = af.Elements<FilterColumn>().FirstOrDefault(fc => fc.ColumnId?.Value == columnId);
                existingColumn?.Remove();

                var fcNew = new FilterColumn { ColumnId = columnId };
                var filters = new Filters();
                foreach (var v in values.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    if (v == null) continue;
                    filters.Append(new Filter { Val = v });
                }
                fcNew.Append(filters);
                af.Append(fcNew);
                ws.Save();
            });
        }

        // -------- Find/Replace --------

        /// <summary>
        /// Finds the first cell text that contains the specified value. Returns the A1 address or null.
        /// Searches values rendered as text (shared strings, inline strings, numbers as invariant strings).
        /// </summary>
        public string? FindFirst(string text)
        {
            if (string.IsNullOrEmpty(text)) return null;
            var ws = _worksheetPart.Worksheet;
            var sd = ws.GetFirstChild<SheetData>();
            if (sd == null) return null;

            foreach (var row in sd.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var t = GetCellText(cell);
                    if (!string.IsNullOrEmpty(t) && t.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
                        return cell.CellReference?.Value;
                }
            }
            return null;
        }

        /// <summary>
        /// Replaces all occurrences of <paramref name="oldText"/> with <paramref name="newText"/> in string cells.
        /// Returns the number of replacements performed.
        /// </summary>
        public int ReplaceAll(string oldText, string newText)
        {
            if (string.IsNullOrEmpty(oldText)) return 0;
            int count = 0;
            WriteLock(() =>
            {
                var ws = _worksheetPart.Worksheet;
                var sd = ws.GetFirstChild<SheetData>();
                if (sd == null) return;

                foreach (var row in sd.Elements<Row>())
                {
                    foreach (var cell in row.Elements<Cell>())
                    {
                        var current = GetCellText(cell);
                        if (string.IsNullOrEmpty(current)) continue;
                        if (current.IndexOf(oldText, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            var replaced = ReplaceIgnoreCase(current, oldText, newText);
                            // write back
                            var (r, c) = Read.A1.ParseCellRef(cell.CellReference?.Value ?? "");
                            CellValue(r, c, replaced);
                            count++;
                        }
                    }
                }
                ws.Save();
            });
            return count;
        }

        private static string ReplaceIgnoreCase(string input, string oldValue, string newValue)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(oldValue)) return input;
            int prev = 0;
            var sb = new StringBuilder(input.Length);
            while (true)
            {
                int idx = input.IndexOf(oldValue, prev, StringComparison.OrdinalIgnoreCase);
                if (idx < 0) break;
                sb.Append(input, prev, idx - prev);
                sb.Append(newValue);
                prev = idx + oldValue.Length;
            }
            sb.Append(input, prev, input.Length - prev);
            return sb.ToString();
        }

        // -------- Validation --------

        /// <summary>
        /// Applies a list validation to the specified A1 range using explicit items.
        /// </summary>
        public void ValidationList(string a1Range, IEnumerable<string> items, bool allowBlank = true)
        {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            if (items == null) throw new ArgumentNullException(nameof(items));

            var joined = string.Join(",", items.Select(i => i?.Replace("\"", "\"\"") ?? string.Empty));
            var formula = "\"" + joined + "\""; // e.g., "New,Processed,Hold"

            WriteLock(() =>
            {
                var ws = _worksheetPart.Worksheet;
                var dvs = ws.GetFirstChild<DataValidations>();
                if (dvs == null)
                {
                    dvs = new DataValidations();
                    ws.Append(dvs);
                }

                var dv = new DataValidation
                {
                    Type = DataValidationValues.List,
                    AllowBlank = allowBlank,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = a1Range }
                };
                dv.Append(new Formula1(formula));
                dvs.Append(dv);
                ws.Save();
            });
        }

        // -------- Sorting (values-only, rewrites range) --------

        /// <summary>
        /// Sorts the sheet's UsedRange rows in-place (excluding header) by the column resolved via header.
        /// Values-only: rewrites cell values; formulas and styles are not preserved.
        /// </summary>
        public void SortUsedRangeByHeader(string header, bool ascending = true)
        {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));

            var a1 = GetUsedRangeA1();
            var (r1, c1, r2, c2) = Read.A1.ParseRange(a1);
            if (r2 - r1 < 1) return; // nothing to sort

            int targetCol = ColumnIndexByHeader(header);
            int targetIndex = targetCol - c1; // zero-based in matrix

            // Read values
            using var rdr = _excelDocument.CreateReader();
            var sh = rdr.GetSheet(this.Name);
            var values = sh.ReadRange(a1);
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            // Build list (rowIndexInSheet, rowValues)
            var list = new List<(int SheetRow, object?[] Row)>();
            for (int r = 1; r < rows; r++) // skip header at 0
            {
                var arr = new object?[cols];
                for (int c = 0; c < cols; c++) arr[c] = values[r, c];
                list.Add((r1 + r, arr));
            }

            int Comparer((int SheetRow, object?[] Row) a, (int SheetRow, object?[] Row) b)
            {
                var va = a.Row[targetIndex];
                var vb = b.Row[targetIndex];
                int cmp = CompareCell(va, vb);
                return ascending ? cmp : -cmp;
            }

            list.Sort(Comparer);

            // Rewrite values back
            WriteLock(() =>
            {
                for (int i = 0; i < list.Count; i++)
                {
                    var row = list[i];
                    int sheetRow = r1 + 1 + i; // first data row
                    for (int c = 0; c < cols; c++)
                    {
                        CellValue(sheetRow, c1 + c, row.Row[c] ?? string.Empty);
                    }
                }
            });

            static int CompareCell(object? a, object? b)
            {
                if (a == null && b == null) return 0;
                if (a == null) return -1;
                if (b == null) return 1;
                if (a is IComparable ca && b is IComparable cb && a.GetType() == b.GetType())
                    return ca.CompareTo(cb);
                // Fallback string compare
                return string.Compare(Convert.ToString(a), Convert.ToString(b), StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
