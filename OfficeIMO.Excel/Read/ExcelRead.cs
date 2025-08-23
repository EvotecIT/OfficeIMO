using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace OfficeIMO.Excel.Read
{
    /// <summary>
    /// Static helpers for quick one‑liner read operations.
    /// These helpers open the workbook read‑only, materialize the requested result, and then dispose it.
    /// Use these when you prefer brevity over streaming/iterator patterns.
    /// </summary>
    public static class ExcelRead
    {
        /// <summary>
        /// Reads a rectangular A1 range (e.g., "A1:C10") from the specified sheet into a dense 2D array of typed values.
        /// </summary>
        /// <param name="path">Path to the .xlsx file.</param>
        /// <param name="sheetName">Worksheet name.</param>
        /// <param name="a1Range">Inclusive A1 range to read.</param>
        /// <param name="options">Optional read options (presets are available in <see cref="ExcelReadPresets"/>).</param>
        /// <returns>Typed matrix with nulls for blank cells.</returns>
        public static object?[,] ReadRange(string path, string sheetName, string a1Range, ExcelReadOptions? options = null)
        {
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            return rdr.GetSheet(sheetName).ReadRange(a1Range);
        }

        /// <summary>
        /// Reads an A1 range into a DataTable. When <paramref name="headersInFirstRow"/> is true, the first row is used as column names.
        /// </summary>
        /// <param name="path">Path to the .xlsx file.</param>
        /// <param name="sheetName">Worksheet name.</param>
        /// <param name="a1Range">Inclusive A1 range to read.</param>
        /// <param name="headersInFirstRow">Whether to treat the first row as column headers.</param>
        /// <param name="options">Optional read options.</param>
        /// <returns>DataTable containing values from the range.</returns>
        public static DataTable ReadRangeAsDataTable(string path, string sheetName, string a1Range, bool headersInFirstRow = true, ExcelReadOptions? options = null)
        {
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            return rdr.GetSheet(sheetName).ReadRangeAsDataTable(a1Range, headersInFirstRow);
        }

        /// <summary>
        /// Reads an A1 range into a dense typed matrix.
        /// </summary>
        /// <typeparam name="T">Element type.</typeparam>
        public static T[,] ReadRangeAs<T>(string path, string sheetName, string a1Range, ExcelReadOptions? options = null)
        {
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            return rdr.GetSheet(sheetName).ReadRangeAs<T>(a1Range);
        }

        /// <summary>
        /// Reads a single‑column A1 range as a typed sequence and materializes it into a list.
        /// </summary>
        /// <typeparam name="T">Target element type.</typeparam>
        /// <param name="path">Path to the .xlsx file.</param>
        /// <param name="sheetName">Worksheet name.</param>
        /// <param name="a1Range">Single‑column A1 range (e.g., "B2:B100").</param>
        /// <param name="options">Optional read options.</param>
        /// <returns>List of typed values.</returns>
        public static List<T> ReadColumnAs<T>(string path, string sheetName, string a1Range, ExcelReadOptions? options = null)
        {
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            return rdr.GetSheet(sheetName).ReadColumnAs<T>(a1Range).ToList();
        }

        /// <summary>
        /// Reads an A1 range into a dense 2D array of typed values by sheet index (1‑based).
        /// </summary>
        /// <param name="path">Path to the .xlsx file.</param>
        /// <param name="sheetIndex">1‑based sheet index in workbook order.</param>
        /// <param name="a1Range">Inclusive A1 range to read.</param>
        /// <param name="options">Optional read options.</param>
        /// <returns>Typed matrix with nulls for blank cells.</returns>
        public static object?[,] ReadRange(string path, int sheetIndex, string a1Range, ExcelReadOptions? options = null)
        {
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            return rdr.GetSheet(sheetIndex).ReadRange(a1Range);
        }

        /// <summary>
        /// Reads an A1 range into a DataTable by sheet index (1‑based).
        /// </summary>
        /// <param name="path">Path to the .xlsx file.</param>
        /// <param name="sheetIndex">1‑based sheet index in workbook order.</param>
        /// <param name="a1Range">Inclusive A1 range to read.</param>
        /// <param name="headersInFirstRow">Whether to treat the first row as column headers.</param>
        /// <param name="options">Optional read options.</param>
        /// <returns>DataTable containing values from the range.</returns>
        public static DataTable ReadRangeAsDataTable(string path, int sheetIndex, string a1Range, bool headersInFirstRow = true, ExcelReadOptions? options = null)
        {
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            return rdr.GetSheet(sheetIndex).ReadRangeAsDataTable(a1Range, headersInFirstRow);
        }

        /// <summary>
        /// Reads an A1 range into a dense typed matrix by sheet index (1‑based).
        /// </summary>
        /// <typeparam name="T">Element type.</typeparam>
        public static T[,] ReadRangeAs<T>(string path, int sheetIndex, string a1Range, ExcelReadOptions? options = null)
        {
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            return rdr.GetSheet(sheetIndex).ReadRangeAs<T>(a1Range);
        }

        /// <summary>
        /// Reads a single‑column A1 range as a typed list by sheet index (1‑based).
        /// </summary>
        /// <typeparam name="T">Target element type.</typeparam>
        /// <param name="path">Path to the .xlsx file.</param>
        /// <param name="sheetIndex">1‑based sheet index in workbook order.</param>
        /// <param name="a1Range">Single‑column A1 range (e.g., "B2:B100").</param>
        /// <param name="options">Optional read options.</param>
        /// <returns>List of typed values.</returns>
        public static List<T> ReadColumnAs<T>(string path, int sheetIndex, string a1Range, ExcelReadOptions? options = null)
        {
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            return rdr.GetSheet(sheetIndex).ReadColumnAs<T>(a1Range).ToList();
        }

        /// <summary>
        /// Reads the used range of a sheet as dictionaries (first row as headers).
        /// </summary>
        public static List<Dictionary<string, object?>> ReadUsedRangeObjects(string path, string sheetName, ExcelReadOptions? options = null)
        {
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            var sh = rdr.GetSheet(sheetName);
            var a1 = sh.GetUsedRangeA1();
            return sh.ReadObjects(a1).ToList();
        }

        /// <summary>
        /// Reads an A1 range as a sequence of dictionaries using the first row as headers.
        /// Keys are header names (case-insensitive in the reader), values are typed when possible.
        /// </summary>
        /// <param name="path">Path to the .xlsx file.</param>
        /// <param name="sheetName">Worksheet name.</param>
        /// <param name="a1Range">Inclusive A1 range (first row must contain headers).</param>
        /// <param name="options">Optional read options.</param>
        /// <returns>Materialized list of row dictionaries.</returns>
        public static List<Dictionary<string, object?>> ReadRangeObjects(string path, string sheetName, string a1Range, ExcelReadOptions? options = null)
        {
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            return rdr.GetSheet(sheetName).ReadObjects(a1Range).ToList();
        }


        /// <summary>
        /// Reads an A1 range from multiple sheets as dictionaries using the first row as headers.
        /// </summary>
        /// <param name="path">Path to the .xlsx file.</param>
        /// <param name="sheetNames">Worksheet names to read.</param>
        /// <param name="a1Range">Inclusive A1 range (first row must contain headers).</param>
        /// <param name="options">Optional read options.</param>
        /// <returns>Dictionary mapping sheet name to list of row dictionaries.</returns>
        public static Dictionary<string, List<Dictionary<string, object?>>> ReadRangeObjectsFromSheets(string path, IEnumerable<string> sheetNames, string a1Range, ExcelReadOptions? options = null)
        {
            var result = new Dictionary<string, List<Dictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            foreach (var name in sheetNames)
            {
                var rows = rdr.GetSheet(name).ReadObjects(a1Range).ToList();
                result[name] = rows;
            }
            return result;
        }

        /// <summary>
        /// Reads an A1 range from every sheet in the workbook as dictionaries using the first row as headers.
        /// </summary>
        public static Dictionary<string, List<Dictionary<string, object?>>> ReadRangeObjectsFromAllSheets(string path, string a1Range, ExcelReadOptions? options = null)
        {
            var result = new Dictionary<string, List<Dictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            foreach (var name in rdr.GetSheetNames())
            {
                var rows = rdr.GetSheet(name).ReadObjects(a1Range).ToList();
                result[name] = rows;
            }
            return result;
        }

        /// <summary>
        /// Reads the used range of the given sheet into a dense 2D array of typed values.
        /// </summary>
        public static object?[,] ReadUsedRange(string path, string sheetName, ExcelReadOptions? options = null)
        {
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            var sh = rdr.GetSheet(sheetName);
            var a1 = sh.GetUsedRangeA1();
            return sh.ReadRange(a1);
        }

        /// <summary>
        /// Reads the used range of the given sheet into a DataTable. First row is used as headers when present.
        /// </summary>
        public static DataTable ReadUsedRangeAsDataTable(string path, string sheetName, bool headersInFirstRow = true, ExcelReadOptions? options = null)
        {
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            var sh = rdr.GetSheet(sheetName);
            var a1 = sh.GetUsedRangeA1();
            return sh.ReadRangeAsDataTable(a1, headersInFirstRow);
        }

        /// <summary>
        /// Reads an A1 range and maps rows (excluding the header row) to instances of <typeparamref name="T"/>.
        /// Header cells are matched to public writable properties on <typeparamref name="T"/> by name (case‑insensitive).
        /// </summary>
        public static System.Collections.Generic.List<T> ReadRangeObjectsAs<T>(string path, string sheetName, string a1Range, ExcelReadOptions? options = null) where T : new()
        {
            using var rdr = ExcelDocumentReader.Open(path, options ?? new ExcelReadOptions());
            return rdr.GetSheet(sheetName).ReadObjects<T>(a1Range).ToList();
        }
    }
}
