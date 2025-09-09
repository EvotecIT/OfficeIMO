using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Read;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Creates or updates a defined name pointing to an A1 range. When <paramref name="scope"/> is provided,
        /// the name is local to that sheet; otherwise it is workbook‑global.
        /// </summary>
        /// <param name="name">Defined name to create or update.</param>
        /// <param name="range">A1 range (e.g. "A1:B10"). Can include a sheet prefix.</param>
        /// <param name="scope">Optional sheet scope for a local name.</param>
        /// <param name="save">When true, saves the workbook after the change.</param>
        /// <param name="hidden">When true, marks the defined name as hidden.</param>
        public void SetNamedRange(string name, string range, ExcelSheet? scope = null, bool save = true, bool hidden = false) {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNullOrWhiteSpace(name);
            ArgumentNullException.ThrowIfNullOrWhiteSpace(range);
#else
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Name cannot be null or whitespace.", nameof(name));
            }
            if (string.IsNullOrWhiteSpace(range)) {
                throw new ArgumentException("Range cannot be null or whitespace.", nameof(range));
            }
#endif

            var workbook = _workBookPart.Workbook;
            var definedNames = workbook.DefinedNames ??= new DefinedNames();

            if (scope == null)
            {
                // Workbook-global name: remove any existing global with same name
                foreach (var dn in definedNames.Elements<DefinedName>().Where(d => d.Name == name && d.LocalSheetId == null).ToList())
                    dn.Remove();
                string reference = NormalizeRange(range); // may already contain a sheet prefix
                var dnNew = new DefinedName { Name = name, Text = reference, Hidden = hidden ? true : (bool?)null };
                definedNames.Append(dnNew);
            }
            else
            {
                // Sheet-local name: remove existing with same name for this sheet
                ushort sheetPos = GetSheetPositionIndex(scope);
                foreach (var dn in definedNames.Elements<DefinedName>().Where(d => d.Name == name && d.LocalSheetId != null && d.LocalSheetId.Value == sheetPos).ToList())
                    dn.Remove();
                string localRef = NormalizeRange(range); // local names store unqualified A1
                var dnNew = new DefinedName { Name = name, Text = localRef, LocalSheetId = sheetPos, Hidden = hidden ? true : (bool?)null };
                definedNames.Append(dnNew);
            }
            if (save) workbook.Save();
        }

        /// <summary>
        /// Sets the print area for a given sheet by creating a sheet-local defined name _xlnm.Print_Area.
        /// </summary>
        public void SetPrintArea(ExcelSheet sheet, string range, bool save = true)
        {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentException("Range cannot be null or whitespace.", nameof(range));

            var workbook = _workBookPart.Workbook;
            var definedNames = workbook.DefinedNames ??= new DefinedNames();

            // Remove existing sheet-local Print_Area for this sheet
            ushort sheetPos = GetSheetPositionIndex(sheet);
            foreach (var dn in definedNames.Elements<DefinedName>().Where(d => d.Name == "_xlnm.Print_Area").ToList())
            {
                if (dn.LocalSheetId != null && dn.LocalSheetId.Value == sheetPos)
                    dn.Remove();
            }

            string normalized = NormalizeRange($"'{sheet.Name}'!{range}");
            var printArea = new DefinedName { Name = "_xlnm.Print_Area", LocalSheetId = sheetPos, Text = normalized };
            definedNames.Append(printArea);
            if (save) workbook.Save();
        }

        /// <summary>
        /// Returns the A1 range for a defined name. If <paramref name="scope"/> is supplied, searches a sheet‑local name first.
        /// </summary>
        /// <param name="name">Defined name to resolve.</param>
        /// <param name="scope">Optional sheet scope to resolve a local name.</param>
        /// <returns>A1 range string or null if not found.</returns>
        public string? GetNamedRange(string name, ExcelSheet? scope = null) {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNullOrWhiteSpace(name);
#else
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Name cannot be null or whitespace.", nameof(name));
            }
#endif
            var definedNames = _workBookPart.Workbook.DefinedNames;
            if (definedNames == null) return null;

            if (scope != null)
            {
                ushort pos = GetSheetPositionIndex(scope);
                var dnLocal = definedNames.Elements<DefinedName>().FirstOrDefault(d => d.Name == name && d.LocalSheetId != null && d.LocalSheetId.Value == pos);
                return dnLocal?.Text;
            }
            else
            {
                var dnGlobal = definedNames.Elements<DefinedName>().FirstOrDefault(d => d.Name == name && d.LocalSheetId == null);
                return dnGlobal?.Text;
            }
        }

        /// <summary>
        /// Returns all defined names with their A1 ranges, optionally limited to a sheet scope.
        /// </summary>
        public IReadOnlyDictionary<string, string> GetAllNamedRanges(ExcelSheet? scope = null) {
            var definedNames = _workBookPart.Workbook.DefinedNames;
            var result = new System.Collections.Generic.Dictionary<string, string>();
            if (definedNames == null) return result;

            if (scope != null)
            {
                ushort pos = GetSheetPositionIndex(scope);
                foreach (var dn in definedNames.Elements<DefinedName>())
                {
                    if (dn.LocalSheetId != null && dn.LocalSheetId.Value == pos)
                        result[dn.Name!] = dn.Text ?? string.Empty;
                }
            }
            else
            {
                foreach (var dn in definedNames.Elements<DefinedName>())
                {
                    if (dn.LocalSheetId == null)
                        result[dn.Name!] = dn.Text ?? string.Empty;
                }
            }
            return result;
        }

        /// <summary>
        /// Removes a defined name. If <paramref name="scope"/> is provided, removes the sheet‑local name; otherwise the global name.
        /// </summary>
        /// <param name="name">Defined name to remove.</param>
        /// <param name="scope">Optional sheet scope.</param>
        /// <param name="save">When true, saves the workbook after removal.</param>
        /// <returns>True if the name existed and was removed; otherwise false.</returns>
        public bool RemoveNamedRange(string name, ExcelSheet? scope = null, bool save = true) {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNullOrWhiteSpace(name);
#else
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Name cannot be null or whitespace.", nameof(name));
            }
#endif
            var definedNames = _workBookPart.Workbook.DefinedNames;
            if (definedNames == null) return false;

            DefinedName? target = null;
            if (scope != null)
            {
                ushort pos = GetSheetPositionIndex(scope);
                target = definedNames.Elements<DefinedName>().FirstOrDefault(d => d.Name == name && d.LocalSheetId != null && d.LocalSheetId.Value == pos);
            }
            else
            {
                target = definedNames.Elements<DefinedName>().FirstOrDefault(d => d.Name == name && d.LocalSheetId == null);
            }
            if (target == null) return false;
            target.Remove();
            if (!definedNames.Elements<DefinedName>().Any()) {
                _workBookPart.Workbook.DefinedNames = null;
            }
            if (save) {
                _workBookPart.Workbook.Save();
            }
            return true;
        }

        private uint GetSheetIndex(ExcelSheet sheet) {
            var sheets = _workBookPart.Workbook.Sheets?.OfType<Sheet>().ToList() ?? new();
            for (int i = 0; i < sheets.Count; i++) {
                if (sheets[i].Name == sheet.Name) {
                    var id = sheets[i].SheetId;
                    if (id == null) {
                        throw new ArgumentException("Worksheet is missing a SheetId.", nameof(sheet));
                    }
                    return id.Value;
                }
            }
            throw new ArgumentException("Worksheet not found in workbook.", nameof(sheet));
        }

        private ushort GetSheetPositionIndex(ExcelSheet sheet)
        {
            var sheets = _workBookPart.Workbook.Sheets?.OfType<Sheet>().ToList() ?? new();
            for (ushort i = 0; i < sheets.Count; i++)
            {
                if (sheets[i].Name == sheet.Name) return i; // 0-based position
            }
            throw new ArgumentException("Worksheet not found in workbook.", nameof(sheet));
        }

        /// <summary>
        /// Normalizes an A1-style range, ensuring absolute references and validating format.
        /// Accepts an optional sheet prefix (e.g. '<c>'Sheet1'!A1:B2</c>').
        /// Throws <see cref="ArgumentException"/> if the input is not a valid A1 range or cell reference.
        /// </summary>
        private static string NormalizeRange(string range) {
            string? sheetPrefix = null;
            string a1 = range;
            int idx = range.IndexOf('!');
            if (idx >= 0) {
                sheetPrefix = range.Substring(0, idx + 1);
                a1 = range.Substring(idx + 1);
            }

            int r1, c1, r2, c2;
            try {
                (r1, c1, r2, c2) = A1.ParseRange(a1);
            } catch (ArgumentException ex) {
                if (a1.Contains(':')) {
                    throw new ArgumentException("Range must be a valid A1 reference such as 'A1:B2'.", nameof(range), ex);
                }
                try {
                    var cell = A1.ParseCellRef(a1);
                    if (cell.Row <= 0 || cell.Col <= 0) {
                        throw new ArgumentException("Range must be a valid A1 reference such as 'A1'.", nameof(range), ex);
                    }
                    r1 = r2 = cell.Row;
                    c1 = c2 = cell.Col;
                } catch (ArgumentException) {
                    throw new ArgumentException("Range must be a valid A1 reference such as 'A1' or 'A1:B2'.", nameof(range), ex);
                }
            }

            string start = $"${A1.ColumnIndexToLetters(c1)}${r1}";
            string end = $"${A1.ColumnIndexToLetters(c2)}${r2}";

            string normalized = start;
            if (start != end) {
                normalized += ":" + end;
            }
            return sheetPrefix + normalized;
        }
    }
}

