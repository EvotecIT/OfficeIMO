using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Write;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>The worksheet tables a native XLSB formula can name, keyed by table name and display name.</summary>
    internal sealed class XlsbTableReferenceIndex {
        internal sealed class Entry {
            internal Entry(uint id, string name, string sheetName, IReadOnlyList<string> columns) {
                Id = id; Name = name; SheetName = sheetName; Columns = columns;
            }
            internal uint Id { get; }
            internal string Name { get; }
            internal string SheetName { get; }
            internal IReadOnlyList<string> Columns { get; }
        }

        private readonly Dictionary<string, Entry> _tables = new Dictionary<string, Entry>(StringComparer.OrdinalIgnoreCase);

        internal static XlsbTableReferenceIndex Create(IEnumerable<ExcelSheet> sheets) {
            var index = new XlsbTableReferenceIndex();
            foreach (ExcelSheet sheet in sheets) {
                foreach (var part in sheet.WorksheetPart.TableDefinitionParts) {
                    Table? table = part.Table;
                    string? display = table?.DisplayName?.Value ?? table?.Name?.Value;
                    if (table?.Id?.Value is not uint id || string.IsNullOrEmpty(display)) continue;
                    string[] columns = table.TableColumns?.Elements<TableColumn>().Select(column => column.Name?.Value ?? string.Empty).ToArray()
                        ?? Array.Empty<string>();
                    var entry = new Entry(id, display!, sheet.Name, columns);
                    index._tables[display!] = entry;
                    if (table.Name?.Value is string name && name.Length != 0) index._tables[name] = entry;
                }
            }
            return index;
        }

        internal bool IsEmpty => _tables.Count == 0;

        /// <summary>True when <paramref name="formula"/> may refer to a table: it has a bracket or contains a table name.</summary>
        internal bool MayReferToTable(string formula) =>
            formula.IndexOf('[') >= 0 || _tables.Keys.Any(name => formula.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0);

        internal bool TryGet(string name, out Entry entry) => _tables.TryGetValue(name, out entry!);
    }

    /// <summary>
    /// Encodes and decodes table structured references as PtgList tokens (MS-XLSB 2.5.98.52). The shared formula parser
    /// has no structured-reference grammar, so each reference is replaced by a placeholder string operand before parsing
    /// and the encoded placeholder token is swapped for the PtgList afterwards.
    /// </summary>
    internal static class XlsbStructuredReferences {
        private const char PlaceholderStart = '\u0001';
        private const char PlaceholderEnd = '\u0002';
        private const byte RowData = 0x00, RowAll = 0x01, RowHeaders = 0x02, RowData2 = 0x04, RowTotals = 0x08, RowCurrent = 0x10;

        /// <summary>Replaces table references in <paramref name="formula"/> with placeholder operands and encodes their PtgList tokens.</summary>
        /// <returns>False with a reason when a reference cannot be encoded; true with the input unchanged when there is none.</returns>
        internal static bool TrySubstitute(string formula, XlsbTableReferenceIndex? tables, LegacyXlsFormulaNameIndex names, int sheetIndex,
            Func<ushort, ushort>? mapExternal, out string rewritten, out IReadOnlyList<byte[]> lists, out string? reason) {
            rewritten = formula;
            lists = Array.Empty<byte[]>();
            reason = null;
            if (tables == null || tables.IsEmpty || !tables.MayReferToTable(formula)) return true;
            ExcelFormulaSyntaxTree tree;
            try {
                tree = ExcelFormulaSyntaxTree.Parse(formula);
            } catch (ArgumentException) {
                return true;
            }
            var encoded = new List<byte[]>();
            string? failure = null;
            string Placeholder(byte[] token) {
                encoded.Add(token);
                return "\"" + PlaceholderStart + (encoded.Count - 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + PlaceholderEnd + "\"";
            }
            string text = tree.RewriteStructuredReferences((tableName, selector) => {
                if (failure != null) return null;
                if (tableName == null) {
                    failure = $"Structured reference '{selector}' needs a table name outside a table definition.";
                    return null;
                }
                if (!tables.TryGet(tableName, out var table)) {
                    failure = $"Table '{tableName}' does not exist in the workbook.";
                    return null;
                }
                if (!TryParseSelector(selector, table, out byte rowType, out byte columns, out ushort first, out ushort last, out string? problem)
                    || !TryCreateList(table, rowType, columns, first, last, names, mapExternal, out byte[] token, out problem)) {
                    failure = problem;
                    return null;
                }
                return Placeholder(token);
            });
            if (failure != null) {
                reason = failure;
                return false;
            }
            // A bare table name stands for the data rows of every column.
            text = ExcelFormulaSyntaxTree.Parse(text).RewriteNames(name => {
                if (failure != null || names.TryGetNameIndex(name, sheetIndex, out _) || !tables.TryGet(name, out var table)) return name;
                if (!TryCreateList(table, RowData, 0, 0, 0, names, mapExternal, out byte[] token, out string? problem)) {
                    failure = problem;
                    return name;
                }
                return Placeholder(token);
            });
            if (failure != null) {
                reason = failure;
                return false;
            }
            rewritten = text;
            lists = encoded;
            return true;
        }

        /// <summary>The PtgList tokens of one formula, swapped for their placeholder string tokens as each leaf is encoded.</summary>
        internal sealed class Placeholders {
            private readonly IReadOnlyList<byte[]> _lists;
            private readonly HashSet<int> _used = new HashSet<int>();

            internal Placeholders(IReadOnlyList<byte[]> lists) => _lists = lists;

            internal bool IsComplete => _used.Count == _lists.Count;

            internal bool TryReplace(byte[] tokens, out byte[] result) {
                result = tokens;
                var spans = new List<(int Offset, int Length, int List)>();
                bool valid = true;
                XlsbFormulaTokenScanner.Visit(tokens, 0, tokens.Length, (offset, token, operand) => {
                    if (token != 0x17) return;
                    int count = tokens[operand] | tokens[operand + 1] << 8;
                    string value = Encoding.Unicode.GetString(tokens, operand + 2, count * 2);
                    if (value.Length < 3 || value[0] != PlaceholderStart || value[value.Length - 1] != PlaceholderEnd) return;
                    if (!int.TryParse(value.Substring(1, value.Length - 2), System.Globalization.NumberStyles.None,
                            System.Globalization.CultureInfo.InvariantCulture, out int list) || list >= _lists.Count || !_used.Add(list)) {
                        valid = false;
                        return;
                    }
                    spans.Add((offset, 3 + count * 2, list));
                });
                if (!valid) return false;
                if (spans.Count == 0) return true;
                using var output = new MemoryStream(tokens.Length + spans.Count * 14);
                int position = 0;
                foreach (var span in spans) {
                    output.Write(tokens, position, span.Offset - position);
                    output.Write(_lists[span.List], 0, _lists[span.List].Length);
                    position = span.Offset + span.Length;
                }
                output.Write(tokens, position, tokens.Length - position);
                result = output.ToArray();
                return true;
            }
        }

        /// <summary>Sets the PtgList data type (bits 10-11 of its flags) from an operand class byte (0x20, 0x40, 0x60).</summary>
        internal static void ApplyOperandClass(byte[] tokens, int offset, byte operandClass) {
            int type = (operandClass >> 5) - 1;
            tokens[offset + 5] = (byte)((tokens[offset + 5] & ~0x0C) | (type << 2));
        }

        /// <summary>Formats the PtgList that starts with eptg at <paramref name="offset"/> and advances past it.</summary>
        internal static bool TryFormat(byte[] tokens, ref int offset, XlsbWorkbook? workbook, out string? text) {
            text = null;
            if (offset > tokens.Length - 13 || tokens[offset] != 0x19) return false;
            ushort flags = (ushort)(tokens[offset + 3] | tokens[offset + 4] << 8);
            uint listIndex = BitConverter.ToUInt32(tokens, offset + 5);
            ushort first = (ushort)(tokens[offset + 9] | tokens[offset + 10] << 8);
            ushort last = (ushort)(tokens[offset + 11] | tokens[offset + 12] << 8);
            offset += 13;
            int columns = flags & 0x03;
            byte rowType = (byte)((flags >> 2) & 0x1F);
            bool invalid = (flags & 0x1000) != 0, nonResident = (flags & 0x2000) != 0;
            if (invalid || nonResident || workbook == null || !workbook.Tables.TryGetValue(listIndex, out XlsbTable? table)) return false;
            if (columns == 3 || (columns != 0 && (first > last || last >= table.Columns.Count)) || (columns == 1 && first != last)) return false;

            string? rows = rowType switch {
                RowData => null,
                RowAll => "#All",
                RowHeaders => "#Headers",
                RowData2 => "#Data",
                0x06 => "#Headers],[#Data",
                RowTotals => "#Totals",
                0x0C => "#Data],[#Totals",
                RowCurrent => "@",
                _ => "?"
            };
            if (rows == "?") return false;
            string? columnText = columns == 0 ? null
                : columns == 1 ? "[" + EncodeColumn(table.Columns[first]) + "]"
                : "[" + EncodeColumn(table.Columns[first]) + "]:[" + EncodeColumn(table.Columns[last]) + "]";
            var builder = new StringBuilder(table.DisplayName);
            if (rows == null) {
                // Excel shows a reference to every data row and column as the bare table name.
                if (columnText != null) {
                    if (columns == 1 && !NeedsBrackets(table.Columns[first])) builder.Append(columnText);
                    else builder.Append('[').Append(columnText).Append(']');
                }
            } else if (rows == "@") {
                if (columnText == null) builder.Append("[@]");
                else if (columns == 1 && !NeedsBrackets(table.Columns[first])) builder.Append("[@").Append(EncodeColumn(table.Columns[first])).Append(']');
                else builder.Append("[@").Append(columnText).Append(']');
            } else if (columnText == null) {
                // One special item keeps one bracket pair; two items list each in its own pair.
                builder.Append(rows.IndexOf(']') >= 0 ? "[[" + rows + "]]" : "[" + rows + "]");
            } else {
                builder.Append("[[").Append(rows).Append("],").Append(columnText).Append(']');
            }
            text = builder.ToString();
            return true;
        }

        private static bool TryCreateList(XlsbTableReferenceIndex.Entry table, byte rowType, byte columns, ushort first, ushort last,
            LegacyXlsFormulaNameIndex names, Func<ushort, ushort>? mapExternal, out byte[] token, out string? reason) {
            token = Array.Empty<byte>();
            reason = null;
            if (!names.TryGetExternSheetIndex(table.SheetName, out ushort generated)) {
                reason = $"The worksheet of table '{table.Name}' has no reference slot.";
                return false;
            }
            ushort ixti = mapExternal?.Invoke(generated) ?? generated;
            token = new byte[14];
            token[0] = 0x18;
            token[1] = 0x19;
            token[2] = (byte)ixti;
            token[3] = (byte)(ixti >> 8);
            // columns in bits 0-1, rowType in bits 2-6; the data type (bits 10-11) is set with the other operand classes.
            ushort flags = (ushort)(columns | rowType << 2);
            token[4] = (byte)flags;
            token[5] = (byte)(flags >> 8);
            BitConverter.GetBytes(table.Id).CopyTo(token, 6);
            token[10] = (byte)first;
            token[11] = (byte)(first >> 8);
            token[12] = (byte)last;
            token[13] = (byte)(last >> 8);
            return true;
        }

        /// <summary>Parses the bracketed selector of a structured reference into PtgList rows and columns.</summary>
        private static bool TryParseSelector(string selector, XlsbTableReferenceIndex.Entry table, out byte rowType, out byte columns,
            out ushort first, out ushort last, out string? reason) {
            rowType = RowData; columns = 0; first = 0; last = 0; reason = null;
            if (selector.Length < 2 || selector[0] != '[' || selector[selector.Length - 1] != ']') {
                reason = $"Structured reference '{table.Name}{selector}' is malformed.";
                return false;
            }
            string inner = selector.Substring(1, selector.Length - 2).Trim();
            bool current = false;
            if (inner.StartsWith("@", StringComparison.Ordinal)) {
                current = true;
                inner = inner.Substring(1).Trim();
            }
            var items = new List<string>();
            var ranges = new List<(string First, string Last)>();
            if (inner.Length == 0) {
                // Table[] or Table[@]: every column.
            } else if (inner[0] != '[') {
                items.Add(inner);
            } else if (!TrySplitItems(inner, items, ranges)) {
                reason = $"Structured reference '{table.Name}{selector}' is malformed.";
                return false;
            }

            var specials = new List<string>();
            var columnItems = new List<string>();
            foreach (string item in items) {
                if (item.StartsWith("#", StringComparison.Ordinal)) specials.Add(item.ToUpperInvariant());
                else columnItems.Add(item);
            }
            if (current && specials.Count != 0) {
                reason = $"Structured reference '{table.Name}{selector}' mixes a row context with special items.";
                return false;
            }
            if (current) rowType = RowCurrent;
            else if (specials.Count != 0) {
                bool all = specials.Contains("#ALL"), headers = specials.Contains("#HEADERS"), data = specials.Contains("#DATA"),
                    totals = specials.Contains("#TOTALS"), thisRow = specials.Contains("#THIS ROW");
                if (specials.Distinct().Count() != specials.Count || specials.Any(item => item is not ("#ALL" or "#HEADERS" or "#DATA" or "#TOTALS" or "#THIS ROW"))) {
                    reason = $"Structured reference '{table.Name}{selector}' uses an unsupported special item.";
                    return false;
                }
                if (all && specials.Count == 1) rowType = RowAll;
                else if (thisRow && specials.Count == 1) rowType = RowCurrent;
                else if (headers && data && specials.Count == 2) rowType = 0x06;
                else if (data && totals && specials.Count == 2) rowType = 0x0C;
                else if (specials.Count == 1) rowType = headers ? RowHeaders : data ? RowData2 : RowTotals;
                else {
                    reason = $"Structured reference '{table.Name}{selector}' combines special items Excel cannot store.";
                    return false;
                }
            }

            if (columnItems.Count + ranges.Count > 1) {
                reason = $"Structured reference '{table.Name}{selector}' names more than one column area.";
                return false;
            }
            if (columnItems.Count == 1) {
                if (!TryFindColumn(table, columnItems[0], out first)) {
                    reason = $"Table '{table.Name}' has no column '{columnItems[0]}'.";
                    return false;
                }
                last = first;
                columns = 1;
            } else if (ranges.Count == 1) {
                if (!TryFindColumn(table, ranges[0].First, out first) || !TryFindColumn(table, ranges[0].Last, out last)) {
                    reason = $"Structured reference '{table.Name}{selector}' names a missing column.";
                    return false;
                }
                if (first > last) (first, last) = (last, first);
                columns = first == last ? (byte)1 : (byte)2;
            }
            return true;
        }

        // Items are bracketed and separated by commas; two column items joined by ':' form a range.
        private static bool TrySplitItems(string inner, List<string> items, List<(string, string)> ranges) {
            var parts = new List<string>();
            var separators = new List<char>();
            int index = 0;
            while (index < inner.Length) {
                while (index < inner.Length && char.IsWhiteSpace(inner[index])) index++;
                if (index >= inner.Length || inner[index] != '[') return false;
                int end = index + 1;
                while (end < inner.Length && !(inner[end] == ']' && !ExcelFormulaSyntaxTree.IsEscapedStructuredCharacter(inner, end))) {
                    if (inner[end] == '[' && !ExcelFormulaSyntaxTree.IsEscapedStructuredCharacter(inner, end)) return false;
                    end++;
                }
                if (end >= inner.Length) return false;
                parts.Add(inner.Substring(index + 1, end - index - 1));
                index = end + 1;
                while (index < inner.Length && char.IsWhiteSpace(inner[index])) index++;
                if (index < inner.Length) {
                    if (inner[index] is not (',' or ':')) return false;
                    separators.Add(inner[index]);
                    index++;
                }
            }
            if (parts.Count == 0 || separators.Count != parts.Count - 1) return false;
            for (int part = 0; part < parts.Count; part++) {
                if (part < separators.Count && separators[part] == ':') {
                    if (parts[part].StartsWith("#", StringComparison.Ordinal) || parts[part + 1].StartsWith("#", StringComparison.Ordinal)
                        || (part + 1 < separators.Count && separators[part + 1] == ':')) return false;
                    ranges.Add((parts[part], parts[part + 1]));
                    part++;
                } else {
                    items.Add(parts[part]);
                }
            }
            return true;
        }

        private static bool TryFindColumn(XlsbTableReferenceIndex.Entry table, string raw, out ushort index) {
            string name = ExcelFormulaSyntaxTree.DecodeStructuredColumnName(raw.Trim());
            for (int column = 0; column < table.Columns.Count; column++) {
                if (string.Equals(table.Columns[column], name, StringComparison.OrdinalIgnoreCase)) {
                    index = (ushort)column;
                    return true;
                }
            }
            index = 0;
            return false;
        }

        private static string EncodeColumn(string name) {
            var builder = new StringBuilder(name.Length);
            foreach (char character in name) {
                if (character is '\'' or '[' or ']' or '#' or '@') builder.Append('\'');
                builder.Append(character);
            }
            return builder.ToString();
        }

        // Excel keeps a single column name in one bracket pair unless it holds a separator or other special character.
        private static bool NeedsBrackets(string name) =>
            name.Length == 0 || char.IsWhiteSpace(name[0]) || char.IsWhiteSpace(name[name.Length - 1])
            || name.IndexOfAny(new[] { '\t', '\n', '\r', ',', ':', '.', '[', ']', '#', '\'', '"', '{', '}', '$', '^', '&', '*', '+', '=', '-', '>', '<', '/', '@' }) >= 0;
    }
}
