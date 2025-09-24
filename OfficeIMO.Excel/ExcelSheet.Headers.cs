using System;
using System.Collections.Generic;
using OfficeIMO.Excel;

namespace OfficeIMO.Excel
{
    /// <summary>
    /// Header-based helpers for addressing cells and columns by header name.
    /// </summary>
    public partial class ExcelSheet
    {
        private Dictionary<string, int>? _headerMapCache;
        private string? _headerMapSourceA1;
        private bool _headerMapNormalize;
        private readonly object _headerMapLock = new object();

        /// <summary>
        /// Builds or returns a cached case-insensitive map of header name to 1-based column index using the first row of UsedRange.
        /// Cache is keyed by UsedRange A1 and NormalizeHeaders option.
        /// </summary>
        public Dictionary<string, int> GetHeaderMap(ExcelReadOptions? options = null)
        {
            var opt = options ?? new ExcelReadOptions();
            var a1Used = GetUsedRangeA1();
            lock (_headerMapLock)
            {
                if (_headerMapCache != null && string.Equals(_headerMapSourceA1, a1Used, StringComparison.Ordinal) && _headerMapNormalize == opt.NormalizeHeaders)
                {
                    return new Dictionary<string, int>(_headerMapCache, StringComparer.OrdinalIgnoreCase);
                }
                using var rdr = _excelDocument.CreateReader(opt);
                var sh = rdr.GetSheet(this.Name);
                var values = sh.ReadRange(a1Used);

                int rows = values.GetLength(0);
                int cols = values.GetLength(1);
                var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                if (rows == 0 || cols == 0)
                {
                    _headerMapCache = map;
                    _headerMapSourceA1 = a1Used;
                    _headerMapNormalize = opt.NormalizeHeaders;
                    return new Dictionary<string, int>(_headerMapCache, StringComparer.OrdinalIgnoreCase);
                }

                var (_, c1, _, _) = A1.ParseRange(a1Used);
                var headers = new string?[cols];
                bool anyHeader = false;
                for (int c = 0; c < cols; c++)
                {
                    var hdr = values[0, c]?.ToString();
                    if (opt.NormalizeHeaders)
                        hdr = System.Text.RegularExpressions.Regex.Replace(hdr ?? string.Empty, "\\s+", " ").Trim();
                    headers[c] = hdr;
                    if (!string.IsNullOrEmpty(hdr))
                        anyHeader = true;
                }

                if (!anyHeader)
                {
                    _headerMapCache = map;
                    _headerMapSourceA1 = a1Used;
                    _headerMapNormalize = opt.NormalizeHeaders;
                    return new Dictionary<string, int>(_headerMapCache, StringComparer.OrdinalIgnoreCase);
                }

                for (int c = 0; c < cols; c++)
                {
                    var hdr = headers[c];
                    if (string.IsNullOrEmpty(hdr))
                        hdr = $"Column{c + 1}";
                    map[hdr] = c1 + c;
                }
                _headerMapCache = map;
                _headerMapSourceA1 = a1Used;
                _headerMapNormalize = opt.NormalizeHeaders;
                return new Dictionary<string, int>(_headerMapCache, StringComparer.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// Returns a 1-based column index for a given header; throws when not found.
        /// </summary>
        public int ColumnIndexByHeader(string header, ExcelReadOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            var map = GetHeaderMap(options);
            if (!map.TryGetValue(header, out var idx))
                throw new KeyNotFoundException($"Header '{header}' not found.");
            return idx;
        }

        /// <summary>
        /// Tries to resolve a 1-based column index for a given header.
        /// </summary>
        public bool TryGetColumnIndexByHeader(string header, out int columnIndex, ExcelReadOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(header))
            {
                columnIndex = 0;
                return false;
            }
            var map = GetHeaderMap(options);
            return map.TryGetValue(header, out columnIndex);
        }

        /// <summary>
        /// Sets a cell value in the specified row by resolving the column using the header name.
        /// </summary>
        public void SetByHeader(int rowIndex, string header, object? value, ExcelReadOptions? options = null)
        {
            if (rowIndex <= 0) throw new ArgumentOutOfRangeException(nameof(rowIndex));
            int col = ColumnIndexByHeader(header, options);
            if (value is null)
                CellValue(rowIndex, col, string.Empty);
            else
                CellValue(rowIndex, col, value);
        }

        /// <summary>
        /// Clears the cached header map.
        /// </summary>
        public void ClearHeaderCache()
        {
            lock (_headerMapLock)
            {
                _headerMapCache = null;
                _headerMapSourceA1 = null;
            }
        }

        /// <summary>
        /// Forces rebuilding the header map for the current UsedRange and options.
        /// </summary>
        public void RefreshHeaderCache(ExcelReadOptions? options = null)
        {
            var _ = GetHeaderMap(options);
        }
    }
}
