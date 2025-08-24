using System;
using System.Collections.Generic;

namespace OfficeIMO.Excel.Read.Edit
{
    /// <summary>
    /// Editable row view over a worksheet range row. Provides header-aware access and setters.
    /// </summary>
    public sealed class RowEdit
    {
        private readonly ExcelSheet _sheet;
        private readonly string[] _headers;
        private readonly Dictionary<string, int> _headerMap;

        internal RowEdit(ExcelSheet sheet, int rowIndex, string[] headers, Dictionary<string, int> headerMap, CellEdit[] cells)
        {
            _sheet = sheet;
            RowIndex = rowIndex;
            _headers = headers;
            _headerMap = headerMap;
            Cells = cells;
        }

        /// <summary>
        /// 1-based row index in the worksheet.
        /// </summary>
        public int RowIndex { get; }

        /// <summary>
        /// All cell handles for this row (1-based indexer via this[int]).
        /// </summary>
        public IReadOnlyList<CellEdit> Cells { get; }

        /// <summary>
        /// 1-based numeric indexer for cells within the materialized range.
        /// </summary>
        public CellEdit this[int colIndex]
        {
            get
            {
                if (colIndex <= 0 || colIndex > Cells.Count) throw new ArgumentOutOfRangeException(nameof(colIndex));
                return Cells[colIndex - 1];
            }
        }

        /// <summary>
        /// Header-aware indexer for cells.
        /// </summary>
        public CellEdit this[string header]
        {
            get
            {
                if (!_headerMap.TryGetValue(header, out var idx))
                    throw new KeyNotFoundException($"Header '{header}' not found.");
                return Cells[idx];
            }
        }

        /// <summary>
        /// Gets the typed value by header.
        /// </summary>
        public T Get<T>(string header)
        {
            var cell = this[header];
            if (cell.Value is null) return default!;
            return cell.ConvertTo<T>();
        }

        /// <summary>
        /// Gets the typed value by header or returns a default.
        /// </summary>
        public T GetOrDefault<T>(string header, T @default = default!)
        {
            var cell = this[header];
            if (cell.Value is null) return @default;
            try { return cell.ConvertTo<T>(); } catch { return @default; }
        }

        /// <summary>
        /// Sets a value by header (writes directly to the worksheet cell).
        /// </summary>
        public void Set(string header, object? value)
        {
            if (!_headerMap.TryGetValue(header, out var idx))
                throw new KeyNotFoundException($"Header '{header}' not found.");
            Cells[idx].Value = value;
        }

        /// <summary>
        /// Returns a cell handle for the given header.
        /// </summary>
        public CellEdit CellByHeader(string header) => this[header];
    }
}

