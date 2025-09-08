using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Read;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        public void SetNamedRange(string name, string range, ExcelSheet? scope = null, bool save = true) {
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

            uint? sheetIndex = scope != null ? GetSheetIndex(scope) : null;

            var existing = definedNames.Elements<DefinedName>().FirstOrDefault(d =>
                d.Name == name && ((sheetIndex == null && d.LocalSheetId == null) ||
                (sheetIndex != null && d.LocalSheetId != null && d.LocalSheetId.Value == sheetIndex)));

            existing?.Remove();

            string reference = scope != null ? $"'{scope.Name}'!{range}" : range;
            reference = NormalizeRange(reference);

            DefinedName dn = new DefinedName {
                Name = name,
                Text = reference
            };
            if (sheetIndex != null) {
                dn.LocalSheetId = sheetIndex;
            }
            definedNames.Append(dn);
            if (save) {
                workbook.Save();
            }
        }

        public string? GetNamedRange(string name, ExcelSheet? scope = null) {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNullOrWhiteSpace(name);
#else
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Name cannot be null or whitespace.", nameof(name));
            }
#endif
            var definedNames = _workBookPart.Workbook.DefinedNames;
            if (definedNames == null) {
                return null;
            }

            uint? sheetIndex = scope != null ? GetSheetIndex(scope) : null;

            var dn = definedNames.Elements<DefinedName>().FirstOrDefault(d =>
                d.Name == name && ((sheetIndex == null && d.LocalSheetId == null) ||
                (sheetIndex != null && d.LocalSheetId != null && d.LocalSheetId.Value == sheetIndex)));

            if (dn == null) {
                return null;
            }

            if (scope != null) {
                string text = dn.Text ?? string.Empty;
                int idx = text.IndexOf('!');
                if (idx >= 0 && idx < text.Length - 1) {
                    return text.Substring(idx + 1);
                }
            }
            return dn.Text;
        }

        public IReadOnlyDictionary<string, string> GetAllNamedRanges(ExcelSheet? scope = null) {
            var definedNames = _workBookPart.Workbook.DefinedNames;
            if (definedNames == null) {
                return new System.Collections.Generic.Dictionary<string, string>();
            }

            uint? sheetIndex = scope != null ? GetSheetIndex(scope) : null;
            var result = new System.Collections.Generic.Dictionary<string, string>();

            foreach (var dn in definedNames.Elements<DefinedName>()) {
                if (sheetIndex == null && dn.LocalSheetId != null) {
                    continue;
                }
                if (sheetIndex != null && (dn.LocalSheetId == null || dn.LocalSheetId.Value != sheetIndex)) {
                    continue;
                }

                string text = dn.Text ?? string.Empty;
                if (scope != null) {
                    int idx = text.IndexOf('!');
                    if (idx >= 0 && idx < text.Length - 1) {
                        text = text.Substring(idx + 1);
                    }
                }
                result[dn.Name!] = text;
            }

            return result;
        }

        public bool RemoveNamedRange(string name, ExcelSheet? scope = null, bool save = true) {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNullOrWhiteSpace(name);
#else
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Name cannot be null or whitespace.", nameof(name));
            }
#endif
            var definedNames = _workBookPart.Workbook.DefinedNames;
            if (definedNames == null) {
                return false;
            }

            uint? sheetIndex = scope != null ? GetSheetIndex(scope) : null;

            var dn = definedNames.Elements<DefinedName>().FirstOrDefault(d =>
                d.Name == name && ((sheetIndex == null && d.LocalSheetId == null) ||
                (sheetIndex != null && d.LocalSheetId != null && d.LocalSheetId.Value == sheetIndex)));

            if (dn == null) {
                return false;
            }

            dn.Remove();
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

