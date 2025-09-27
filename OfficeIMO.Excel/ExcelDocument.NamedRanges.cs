using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static string EscapeSheetName(string name) {
            return (name ?? string.Empty).Replace("'", "''");
        }
        private static string StripSheetPrefixIfMatches(string text, ExcelSheet scope) {
            if (string.IsNullOrEmpty(text)) return text;
            string quoted = "'" + EscapeSheetName(scope.Name) + "'!";
            if (text.StartsWith(quoted, System.StringComparison.Ordinal)) {
                return text.Substring(quoted.Length);
            }
            string unquoted = scope.Name + "!";
            if (text.StartsWith(unquoted, System.StringComparison.Ordinal)) {
                return text.Substring(unquoted.Length);
            }
            return text;
        }
        /// <summary>
        /// Creates or updates a defined name pointing to an A1 range. When <paramref name="scope"/> is provided,
        /// the name is local to that sheet; otherwise it is workbook‑global.
        /// </summary>
        /// <param name="name">Defined name to create or update.</param>
        /// <param name="range">A1 range (e.g. "A1:B10"). Can include a sheet prefix.</param>
        /// <param name="scope">Optional sheet scope for a local name.</param>
        /// <param name="save">When true, saves the workbook after the change.</param>
        /// <param name="hidden">When true, marks the defined name as hidden.</param>
        /// <param name="validationMode">Controls how the name and range are validated: Sanitize (default) clamps/adjusts; Strict throws on invalid input.</param>
        public void SetNamedRange(string name, string range, ExcelSheet? scope = null, bool save = true, bool hidden = false, NameValidationMode validationMode = NameValidationMode.Sanitize) {
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

            // Validate or sanitize the defined name
            name = EnsureValidDefinedName(name, validationMode);

            if (scope == null) {
                // Workbook-global name: remove any existing global with same name
                foreach (var dn in definedNames.Elements<DefinedName>().Where(d => d.Name == name && d.LocalSheetId == null).ToList())
                    dn.Remove();
                string reference = NormalizeRange(range, validationMode); // may already contain a sheet prefix
                var dnNew = new DefinedName { Name = name, Text = reference, Hidden = hidden ? true : (bool?)null };
                definedNames.Append(dnNew);
            } else {
                // Sheet-local name: remove existing with same name for this sheet
                ushort sheetPos = GetSheetPositionIndex(scope);
                foreach (var dn in definedNames.Elements<DefinedName>().Where(d => d.Name == name && d.LocalSheetId != null && d.LocalSheetId.Value == sheetPos).ToList())
                    dn.Remove();
                // Use an explicit sheet-qualified reference for maximum Excel compatibility
                // Escape single quotes inside the sheet name per Excel syntax ('' -> ')
                string sheetQuoted = $"'{EscapeSheetName(scope.Name)}'!";
                string localRef = NormalizeRange(sheetQuoted + range, validationMode);
                var dnNew = new DefinedName { Name = name, Text = localRef, LocalSheetId = sheetPos, Hidden = hidden ? true : (bool?)null };
                definedNames.Append(dnNew);
            }
            if (save) workbook.Save();
        }

        /// <summary>
        /// Sets the print area for a given sheet by creating a sheet-local defined name _xlnm.Print_Area.
        /// </summary>
        public void SetPrintArea(ExcelSheet sheet, string range, bool save = true) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentException("Range cannot be null or whitespace.", nameof(range));

            var workbook = _workBookPart.Workbook;
            var definedNames = workbook.DefinedNames ??= new DefinedNames();

            // Remove existing sheet-local Print_Area for this sheet
            ushort sheetPos = GetSheetPositionIndex(sheet);
            foreach (var dn in definedNames.Elements<DefinedName>().Where(d => d.Name == "_xlnm.Print_Area").ToList()) {
                if (dn.LocalSheetId != null && dn.LocalSheetId.Value == sheetPos)
                    dn.Remove();
            }

            string normalized = NormalizeRange($"'{EscapeSheetName(sheet.Name)}'!{range}");
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

            if (scope != null) {
                ushort pos = GetSheetPositionIndex(scope);
                var dnLocal = definedNames.Elements<DefinedName>().FirstOrDefault(d => d.Name == name && d.LocalSheetId != null && d.LocalSheetId.Value == pos);
                var text = dnLocal?.Text;
                if (string.IsNullOrEmpty(text)) return text;
                return StripSheetPrefixIfMatches(text!, scope);
            } else {
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

            if (scope != null) {
                ushort pos = GetSheetPositionIndex(scope);
                foreach (var dn in definedNames.Elements<DefinedName>()) {
                    if (dn.LocalSheetId != null && dn.LocalSheetId.Value == pos) {
                        var text = dn.Text ?? string.Empty;
                        result[dn.Name!] = StripSheetPrefixIfMatches(text, scope);
                    }
                }
            } else {
                foreach (var dn in definedNames.Elements<DefinedName>()) {
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
            if (scope != null) {
                ushort pos = GetSheetPositionIndex(scope);
                target = definedNames.Elements<DefinedName>().FirstOrDefault(d => d.Name == name && d.LocalSheetId != null && d.LocalSheetId.Value == pos);
            } else {
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

        private ushort GetSheetPositionIndex(ExcelSheet sheet) {
            var sheets = _workBookPart.Workbook.Sheets?.OfType<Sheet>().ToList() ?? new();
            for (ushort i = 0; i < sheets.Count; i++) {
                if (sheets[i].Name == sheet.Name) return i; // 0-based position
            }
            throw new ArgumentException("Worksheet not found in workbook.", nameof(sheet));
        }

        /// <summary>
        /// Sets rows/columns to repeat at top/left when printing a specific sheet by creating a sheet-local
        /// defined name _xlnm.Print_Titles. Pass nulls to clear existing print titles.
        /// </summary>
        /// <param name="sheet">Target sheet.</param>
        /// <param name="firstRow">First row to repeat (1-based), or null.</param>
        /// <param name="lastRow">Last row to repeat (1-based), or null.</param>
        /// <param name="firstCol">First column to repeat (1-based), or null.</param>
        /// <param name="lastCol">Last column to repeat (1-based), or null.</param>
        /// <param name="save">Whether to save the workbook after the change.</param>
        public void SetPrintTitles(ExcelSheet sheet, int? firstRow, int? lastRow, int? firstCol, int? lastCol, bool save = true) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));

            var workbook = _workBookPart.Workbook;
            var definedNames = workbook.DefinedNames ??= new DefinedNames();

            // Remove existing sheet-local Print_Titles for this sheet
            ushort sheetPos = GetSheetPositionIndex(sheet);
            foreach (var dn in definedNames.Elements<DefinedName>().Where(d => d.Name == "_xlnm.Print_Titles").ToList()) {
                if (dn.LocalSheetId != null && dn.LocalSheetId.Value == sheetPos)
                    dn.Remove();
            }

            // Nothing to set? stop here (clears existing titles)
            bool hasRows = firstRow.HasValue && lastRow.HasValue && firstRow.Value > 0 && lastRow.Value >= firstRow.Value;
            bool hasCols = firstCol.HasValue && lastCol.HasValue && firstCol.Value > 0 && lastCol.Value >= firstCol.Value;
            if (!hasRows && !hasCols) {
                if (save) workbook.Save();
                return;
            }

            string? rowsPart = null, colsPart = null;
            if (hasRows) {
                rowsPart = $"'{EscapeSheetName(sheet.Name)}'!${firstRow.GetValueOrDefault()}:${lastRow.GetValueOrDefault()}";
            }
            if (hasCols) {
                string c1 = A1.ColumnIndexToLetters(firstCol.GetValueOrDefault());
                string c2 = A1.ColumnIndexToLetters(lastCol.GetValueOrDefault());
                colsPart = $"'{EscapeSheetName(sheet.Name)}'!${c1}:${c2}";
            }

            string text = hasRows && hasCols ? string.Concat(rowsPart, ",", colsPart) : (rowsPart ?? colsPart)!;
            var dnNew = new DefinedName { Name = "_xlnm.Print_Titles", LocalSheetId = sheetPos, Text = text };
            definedNames.Append(dnNew);
            if (save) workbook.Save();
        }

        /// <summary>
        /// Repairs common issues with defined names that can trigger Excel's file repair, such as
        /// duplicates within the same scope, invalid LocalSheetId after sheet reordering/removal,
        /// or references containing #REF!.
        /// </summary>
        internal void RepairDefinedNames(bool save = true) {
            var wb = _workBookPart.Workbook;
            var definedNames = wb.DefinedNames;
            if (definedNames == null) return;

            var sheets = wb.Sheets?.OfType<Sheet>().ToList() ?? new();
            int sheetCount = sheets.Count;

            var toRemove = new System.Collections.Generic.HashSet<DocumentFormat.OpenXml.Spreadsheet.DefinedName>();
            var seen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var dn in definedNames.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>()) {
                string? name = dn.Name;
                if (string.IsNullOrWhiteSpace(name)) { toRemove.Add(dn); continue; }

                uint? local = dn.LocalSheetId?.Value;
                if (local.HasValue && (local.Value >= (uint)sheetCount)) { toRemove.Add(dn); continue; }

                string key = (local.HasValue ? local.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "G") + "|" + name;
                if (!seen.Add(key)) { toRemove.Add(dn); continue; }

                string text = dn.Text ?? string.Empty;
                if (text.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0) { toRemove.Add(dn); continue; }
            }

            if (toRemove.Count > 0) {
                foreach (var dn in toRemove) dn.Remove();
                if (!definedNames.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>().Any()) {
                    wb.DefinedNames = null;
                }
                if (save) wb.Save();
            }
        }

        /// <summary>
        /// Normalizes an A1-style range, ensuring absolute references and validating format.
        /// Accepts an optional sheet prefix (e.g. '<c>'Sheet1'!A1:B2</c>').
        /// Throws <see cref="ArgumentException"/> if the input is not a valid A1 range or cell reference.
        /// </summary>
        private static string NormalizeRange(string range) {
            return NormalizeRange(range, NameValidationMode.Sanitize);
        }

        private static string NormalizeRange(string range, NameValidationMode validationMode) {
            string? sheetPrefix = null;
            string a1 = range;
            int idx = range.IndexOf('!');
            if (idx >= 0) {
                sheetPrefix = range.Substring(0, idx + 1);
                a1 = range.Substring(idx + 1);
            }

            int r1, c1, r2, c2;
            if (!A1.TryParseRange(a1, out r1, out c1, out r2, out c2)) {
                bool containsColon = a1.IndexOf(':') >= 0;
                var cell = A1.ParseCellRef(a1);
                if (cell.Row <= 0 || cell.Col <= 0) {
                    string message = containsColon
                        ? "Range must be a valid A1 reference such as 'A1:B2'."
                        : "Range must be a valid A1 reference such as 'A1' or 'A1:B2'.";
                    throw new ArgumentException(message, nameof(range));
                }
                r1 = r2 = cell.Row;
                c1 = c2 = cell.Col;
            }

            // Bounds check: Excel supports 1..1,048,576 rows and 1..16,384 columns (XFD)
            const int MaxRow = 1_048_576;
            const int MaxCol = 16_384;
            bool outOfBounds = (r1 < 1 || c1 < 1 || r2 > MaxRow || c2 > MaxCol || c1 > MaxCol || r1 > MaxRow);
            if (outOfBounds && validationMode == NameValidationMode.Strict)
                throw new ArgumentOutOfRangeException(nameof(range), "A1 range exceeds Excel bounds (rows ≤ 1,048,576; cols ≤ 16,384). Use Sanitize to clamp.");
            // Sanitize: clamp into valid range
            r1 = Math.Max(1, Math.Min(MaxRow, r1));
            r2 = Math.Max(1, Math.Min(MaxRow, r2));
            c1 = Math.Max(1, Math.Min(MaxCol, c1));
            c2 = Math.Max(1, Math.Min(MaxCol, c2));

            string start = $"${A1.ColumnIndexToLetters(c1)}${r1}";
            string end = $"${A1.ColumnIndexToLetters(c2)}${r2}";

            string normalized = start;
            if (start != end) {
                normalized += ":" + end;
            }
            return sheetPrefix + normalized;
        }

        /// <summary>
        /// Ensures a defined name complies with Excel rules. In Sanitize mode, returns a corrected name.
        /// Throws in Strict mode when input is invalid.
        /// Rules:
        /// - 1..255 characters
        /// - First char must be a letter or underscore
        /// - Allowed characters: letters, digits, underscore, period
        /// - Cannot look like a cell reference (e.g., A1, AA10) or an R1C1 reference
        /// - Cannot be TRUE or FALSE (case-insensitive)
        /// </summary>
        private static string EnsureValidDefinedName(string name, NameValidationMode mode) {
            const int MaxLen = 255;
            if (string.IsNullOrWhiteSpace(name)) {
                if (mode == NameValidationMode.Strict) throw new System.ArgumentException($"Defined name '{name}' cannot be null or whitespace.", nameof(name));
                name = "_";
            }

            // Trim spaces and replace invalid chars
            var sb = new System.Text.StringBuilder(name.Length);
            foreach (char ch in name.Trim()) {
                if (char.IsLetterOrDigit(ch) || ch == '_' || ch == '.') sb.Append(ch);
                else { sb.Append('_'); }
            }
            if (sb.Length == 0) { sb.Append('_'); }
            if (!char.IsLetter(sb[0]) && sb[0] != '_') { sb.Insert(0, '_'); }

            // Disallow TRUE/FALSE exactly (case-insensitive)
            var normalized = sb.ToString();
            if (string.Equals(normalized, "TRUE", System.StringComparison.OrdinalIgnoreCase) || string.Equals(normalized, "FALSE", System.StringComparison.OrdinalIgnoreCase)) {
                if (mode == NameValidationMode.Strict) throw new System.ArgumentException($"Defined name '{name}' cannot be TRUE or FALSE.", nameof(name));
                normalized = "_" + normalized;
            }

            // Avoid names that look like A1 cell references or R1C1 format
            bool LooksLikeA1(string s) {
                var t = OfficeIMO.Excel.A1.ParseCellRef(s);
                return t.Row > 0 && t.Col > 0;
            }
            bool LooksLikeR1C1(string s) {
                // Very lenient check: R<digits>C<digits>
                if (s.Length < 3) return false;
                if (s[0] != 'R' && s[0] != 'r') return false;
                int i = 1; while (i < s.Length && char.IsDigit(s[i])) i++;
                if (i == 1 || i >= s.Length || (s[i] != 'C' && s[i] != 'c')) return false;
                i++; if (i >= s.Length) return false;
                int j = i; while (j < s.Length && char.IsDigit(s[j])) j++;
                return j > i && j == s.Length;
            }

            if (LooksLikeA1(normalized) || LooksLikeR1C1(normalized)) {
                if (mode == NameValidationMode.Strict) throw new System.ArgumentException($"Defined name '{name}' cannot be a cell address or R1C1 reference.", nameof(name));
                normalized = "_" + normalized;
            }

            if (normalized.Length > MaxLen) {
                if (mode == NameValidationMode.Strict) throw new System.ArgumentException($"Defined name '{name}' exceeds maximum length of {MaxLen} characters (actual {normalized.Length}).", nameof(name));
                normalized = normalized.Substring(0, MaxLen);
            }

            return normalized;
        }
    }
}

