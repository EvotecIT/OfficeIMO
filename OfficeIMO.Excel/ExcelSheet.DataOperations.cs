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
        /// When the header is missing the operation is skipped.
        /// </summary>
        public void AutoFilterByHeaderEquals(string header, IEnumerable<string> values)
        {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            if (values == null) throw new ArgumentNullException(nameof(values));

            WriteLock(() =>
            {
                if (!TryGetColumnIndexByHeader(header, out var colIndex))
                    return;
                var ws = _worksheetPart.Worksheet;
                var af = ws.Elements<AutoFilter>().FirstOrDefault();
                if (af == null)
                {
                    var used = GetUsedRangeA1();
                    af = new AutoFilter { Reference = used };
                    ws.InsertAfter(af, ws.GetFirstChild<SheetData>());
                }

                var (r1, c1, r2, c2) = A1.ParseRange(af.Reference!);
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

        /// <summary>
        /// Applies equals filters for multiple headers at once (AND semantics across columns, OR semantics within a column).
        /// Headers that cannot be resolved are ignored.
        /// </summary>
        public void AutoFilterByHeadersEquals(params (string Header, IEnumerable<string> Values)[] filters)
        {
            if (filters == null || filters.Length == 0) throw new ArgumentException("At least one filter must be provided.", nameof(filters));
            WriteLock(() =>
            {
                var toApply = new List<(int ColumnIndex, IEnumerable<string> Values)>();
                foreach (var (header, values) in filters)
                {
                    if (string.IsNullOrWhiteSpace(header)) continue;
                    if (!TryGetColumnIndexByHeader(header, out var colIndex)) continue;
                    toApply.Add((colIndex, values ?? Array.Empty<string>()));
                }
                if (toApply.Count == 0) return;

                var ws = _worksheetPart.Worksheet;
                var af = ws.Elements<AutoFilter>().FirstOrDefault();
                if (af == null)
                {
                    var used = GetUsedRangeA1();
                    af = new AutoFilter { Reference = used };
                    ws.InsertAfter(af, ws.GetFirstChild<SheetData>());
                }

                var (r1, c1, r2, c2) = A1.ParseRange(af.Reference!);

                foreach (var (colIndex, values) in toApply)
                {
                    if (colIndex < c1 || colIndex > c2) continue;
                    uint columnId = (uint)(colIndex - c1);

                    var existingColumn = af.Elements<FilterColumn>().FirstOrDefault(fc => fc.ColumnId?.Value == columnId);
                    existingColumn?.Remove();

                    var fcNew = new FilterColumn { ColumnId = columnId };
                    var filtersNode = new Filters();
                    foreach (var v in values.Distinct(StringComparer.OrdinalIgnoreCase))
                    {
                        if (v == null) continue;
                        filtersNode.Append(new Filter { Val = v });
                    }
                    fcNew.Append(filtersNode);
                    af.Append(fcNew);
                }
                ws.Save();
            });
        }

        /// <summary>
        /// Applies an AutoFilter text contains filter to a column resolved by header within the current AutoFilter range.
        /// Uses wildcard pattern matching ("*text*") via CustomFilters with Equal operator.
        /// When the header is missing the operation is skipped.
        /// </summary>
        public void AutoFilterByHeaderContains(string header, string containsText)
        {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            if (containsText is null) throw new ArgumentNullException(nameof(containsText));

            WriteLock(() =>
            {
                if (!TryGetColumnIndexByHeader(header, out var colIndex))
                    return;
                var ws = _worksheetPart.Worksheet;
                var af = ws.Elements<AutoFilter>().FirstOrDefault();
                if (af == null)
                {
                    var used = GetUsedRangeA1();
                    af = new AutoFilter { Reference = used };
                    ws.InsertAfter(af, ws.GetFirstChild<SheetData>());
                }

                var (r1, c1, r2, c2) = A1.ParseRange(af.Reference!);
                if (colIndex < c1 || colIndex > c2)
                    throw new ArgumentOutOfRangeException(nameof(header), $"Header '{header}' is outside the AutoFilter range {af.Reference}.");

                uint columnId = (uint)(colIndex - c1);

                var existingColumn = af.Elements<FilterColumn>().FirstOrDefault(fc => fc.ColumnId?.Value == columnId);
                existingColumn?.Remove();

                var fcNew = new FilterColumn { ColumnId = columnId };
                var custom = new CustomFilters();
                // Excel uses Equal + wildcard pattern for contains
                custom.Append(new CustomFilter { Operator = FilterOperatorValues.Equal, Val = "*" + containsText + "*" });
                fcNew.Append(custom);
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
                            var (r, c) = A1.ParseCellRef(cell.CellReference?.Value ?? "");
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

        /// <summary>
        /// Applies a whole number validation to the specified A1 range.
        /// </summary>
        public void ValidationWholeNumber(string a1Range, DataValidationOperatorValues @operator, int formula1, int? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null)
        {
            string f1 = formula1.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string? f2 = formula2?.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ValidationAdd(a1Range, DataValidationValues.Whole, @operator, f1, f2, allowBlank, errorTitle, errorMessage);
        }

        /// <summary>
        /// Applies a decimal number validation to the specified A1 range.
        /// </summary>
        public void ValidationDecimal(string a1Range, DataValidationOperatorValues @operator, double formula1, double? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null)
        {
            string f1 = formula1.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string? f2 = formula2?.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ValidationAdd(a1Range, DataValidationValues.Decimal, @operator, f1, f2, allowBlank, errorTitle, errorMessage);
        }

        /// <summary>
        /// Applies a date validation to the specified A1 range.
        /// </summary>
        public void ValidationDate(string a1Range, DataValidationOperatorValues @operator, DateTime formula1, DateTime? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null)
        {
            string f1 = formula1.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
            string? f2 = formula2?.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
            ValidationAdd(a1Range, DataValidationValues.Date, @operator, f1, f2, allowBlank, errorTitle, errorMessage);
        }

        /// <summary>
        /// Applies a time validation to the specified A1 range.
        /// </summary>
        public void ValidationTime(string a1Range, DataValidationOperatorValues @operator, TimeSpan formula1, TimeSpan? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null)
        {
            string f1 = formula1.TotalDays.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string? f2 = formula2?.TotalDays.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ValidationAdd(a1Range, DataValidationValues.Time, @operator, f1, f2, allowBlank, errorTitle, errorMessage);
        }

        /// <summary>
        /// Applies a text length validation to the specified A1 range.
        /// </summary>
        public void ValidationTextLength(string a1Range, DataValidationOperatorValues @operator, int formula1, int? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null)
        {
            string f1 = formula1.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string? f2 = formula2?.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ValidationAdd(a1Range, DataValidationValues.TextLength, @operator, f1, f2, allowBlank, errorTitle, errorMessage);
        }

        /// <summary>
        /// Applies a custom formula validation to the specified A1 range.
        /// </summary>
        public void ValidationCustomFormula(string a1Range, string formula, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null)
        {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentNullException(nameof(formula));
            ValidationAdd(a1Range, DataValidationValues.Custom, null, formula, null, allowBlank, errorTitle, errorMessage);
        }

        private void ValidationAdd(string a1Range, DataValidationValues type, DataValidationOperatorValues? @operator, string formula1, string? formula2, bool allowBlank, string? errorTitle, string? errorMessage)
        {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            if (string.IsNullOrWhiteSpace(formula1)) throw new ArgumentNullException(nameof(formula1));

            bool requiresTwo = @operator == DataValidationOperatorValues.Between || @operator == DataValidationOperatorValues.NotBetween;
            if (requiresTwo && formula2 == null) throw new ArgumentNullException(nameof(formula2));

            DataValidation dv = new DataValidation
            {
                Type = type,
                AllowBlank = allowBlank,
                Operator = @operator,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = a1Range }
            };

            if (!string.IsNullOrEmpty(errorTitle) || !string.IsNullOrEmpty(errorMessage))
            {
                dv.ShowErrorMessage = true;
                dv.ErrorTitle = errorTitle;
                dv.Error = errorMessage;
            }

            dv.Append(new Formula1(formula1));
            if (formula2 != null)
            {
                dv.Append(new Formula2(formula2));
            }

            WriteLock(() =>
            {
                Worksheet ws = _worksheetPart.Worksheet;
                DataValidations? dvs = ws.GetFirstChild<DataValidations>();
                if (dvs == null)
                {
                    dvs = new DataValidations();
                    ws.Append(dvs);
                }
                dvs.Append(dv);
                ws.Save();
            });
        }

        // -------- Sorting (values-only, rewrites range) --------

        /// <summary>
        /// Sorts the sheet's UsedRange rows in-place (excluding header) by the column resolved via header.
        /// Values-only: rewrites cell values; formulas and styles are not preserved.
        /// When the header cannot be resolved the sort is skipped.
        /// </summary>
        public void SortUsedRangeByHeader(string header, bool ascending = true)
        {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));

            var a1 = GetUsedRangeA1();
            var (r1, c1, r2, c2) = A1.ParseRange(a1);
            if (r2 - r1 < 1) return; // nothing to sort

            if (!TryGetColumnIndexByHeader(header, out var targetCol)) return;
            if (targetCol < c1 || targetCol > c2) return;
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

        /// <summary>
        /// Sorts the sheet's UsedRange rows in-place by multiple headers in the given order (excluding header row).
        /// Values-only: rewrites cell values; formulas and styles are not preserved. Missing headers are ignored and when no
        /// valid headers remain the sort is skipped.
        /// </summary>
        public void SortUsedRangeByHeaders(params (string Header, bool Ascending)[] keys)
        {
            if (keys == null || keys.Length == 0) throw new ArgumentException("At least one key is required.", nameof(keys));

            var a1 = GetUsedRangeA1();
            var (r1, c1, r2, c2) = A1.ParseRange(a1);
            if (r2 - r1 < 1) return;

            // Resolve column indices and validate
            var cols = new List<(int Index, bool Asc)>();
            foreach (var k in keys)
            {
                if (string.IsNullOrWhiteSpace(k.Header)) continue;
                if (!TryGetColumnIndexByHeader(k.Header, out var col)) continue;
                if (col < c1 || col > c2) continue;
                cols.Add((col - c1, k.Ascending));
            }
            if (cols.Count == 0) return;

            using var rdr = _excelDocument.CreateReader();
            var sh = rdr.GetSheet(this.Name);
            var values = sh.ReadRange(a1);
            int rows = values.GetLength(0);
            int width = values.GetLength(1);

            var list = new List<(int SheetRow, object?[] Row)>();
            for (int r = 1; r < rows; r++)
            {
                var arr = new object?[width];
                for (int c = 0; c < width; c++) arr[c] = values[r, c];
                list.Add((r1 + r, arr));
            }

            list.Sort((a, b) =>
            {
                foreach (var (idx, asc) in cols)
                {
                    int cmp = CompareCell(a.Row[idx], b.Row[idx]);
                    if (cmp != 0) return asc ? cmp : -cmp;
                }
                return 0;
            });

            WriteLock(() =>
            {
                for (int i = 0; i < list.Count; i++)
                {
                    var row = list[i];
                    int sheetRow = r1 + 1 + i;
                    for (int c = 0; c < width; c++)
                        CellValue(sheetRow, c1 + c, row.Row[c] ?? string.Empty);
                }
            });

            static int CompareCell(object? a, object? b)
            {
                if (a == null && b == null) return 0;
                if (a == null) return -1;
                if (b == null) return 1;
                if (a is IComparable ca && b is IComparable cb && a.GetType() == b.GetType())
                    return ca.CompareTo(cb);
                return string.Compare(Convert.ToString(a), Convert.ToString(b), StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
