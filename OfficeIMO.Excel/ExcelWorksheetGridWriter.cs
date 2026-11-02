using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Serializes a loaded worksheet with the bytes the SDK's writer produces, writing the cell grid directly because
    /// the SDK's element writer costs most of saving an edited workbook. The SDK still writes everything around the grid.
    /// </summary>
    /// <remarks>
    /// Only rows of cells holding an optional formula and value are written here, in namespaces the root declares, with
    /// text that needs no XML writer decisions. Anything else returns false before the output is used, and the caller
    /// saves the root through the SDK. <see cref="ExcelWorksheetGridLoader"/> reads the same shapes.
    /// </remarks>
    internal static class ExcelWorksheetGridWriter {
        private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        private static readonly byte[] EmptyGridSuffix = Encoding.ASCII.GetBytes("sheetData />");

        /// <summary>Writes <paramref name="worksheet"/> to <paramref name="output"/> as its Save would, or returns false with the output unspecified.</summary>
        internal static bool TrySave(Worksheet worksheet, ExcelPartBuffer output) {
            SheetData? sheetData = null;
            for (OpenXmlElement? child = worksheet.FirstChild; child != null; child = child.NextSibling()) {
                if (child is not SheetData data) continue;
                if (sheetData != null) return false;
                sheetData = data;
            }
            if (sheetData == null || sheetData.FirstChild == null || !IsBare(sheetData)) return false;

            var scope = new Scope(worksheet);
            if (!scope.IsMainPrefix(sheetData.Prefix)) return false;

            var grid = new ExcelPartBuffer();
            var empty = new SheetData();
            bool swapped = false;
            try {
                var writer = new GridWriter(grid, scope, sheetData.Prefix);
                if (!writer.TryWriteRows(sheetData)) return false;

                worksheet.ReplaceChild(empty, sheetData);
                swapped = true;
                output.Clear();
                worksheet.Save(output);
                worksheet.ReplaceChild(sheetData, empty);
                swapped = false;

                if (!TryFindEmptyGrid(output, out int start, out int end)) return false;
                int tailLength = output.ContentLength - end;
                byte[] tail = new byte[tailLength];
                Buffer.BlockCopy(output.Buffer, end, tail, 0, tailLength);
                output.Truncate(start);
                string tag = Qualified(sheetData.Prefix, "sheetData");
                output.WriteAscii("<" + tag + ">");
                output.Write(grid.Buffer, 0, grid.ContentLength);
                output.WriteAscii("</" + tag + ">");
                output.Write(tail, 0, tail.Length);
                return true;
            } finally {
                if (swapped) worksheet.ReplaceChild(sheetData, empty);
                grid.Release();
            }
        }

        /// <summary>Finds the single empty grid the SDK wrote in place of the real one.</summary>
        private static bool TryFindEmptyGrid(ExcelPartBuffer output, out int start, out int end) {
            start = end = -1;
            ReadOnlySpan<byte> xml = new ReadOnlySpan<byte>(output.Buffer, 0, output.ContentLength);
            int suffix = xml.IndexOf(EmptyGridSuffix);
            if (suffix < 0 || xml.Slice(suffix + EmptyGridSuffix.Length).IndexOf(EmptyGridSuffix) >= 0) return false;
            int open = xml.Slice(0, suffix).LastIndexOf((byte)'<');
            if (open < 0) return false;
            foreach (byte character in xml.Slice(open + 1, suffix - open - 1)) {
                if (character == (byte)' ' || character == (byte)'"') return false;
            }
            start = open;
            end = suffix + EmptyGridSuffix.Length;
            return true;
        }

        private static bool IsBare(OpenXmlElement element) =>
            !element.HasAttributes && element.MCAttributes == null && !element.NamespaceDeclarations.Any();

        private static string Qualified(string prefix, string localName) =>
            prefix.Length == 0 ? localName : prefix + ":" + localName;

        /// <summary>The prefixes the root binds, which are the only ones the grid may use without a declaration of its own.</summary>
        private sealed class Scope {
            private const string DyDescentNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
            private readonly Dictionary<string, string> _namespaces = new(StringComparer.Ordinal);

            internal Scope(Worksheet worksheet) {
                foreach (KeyValuePair<string, string> declaration in worksheet.NamespaceDeclarations) {
                    _namespaces[declaration.Key] = declaration.Value;
                    if (string.Equals(declaration.Value, DyDescentNamespace, StringComparison.Ordinal) && declaration.Key.Length > 0) {
                        DyDescentPrefix ??= declaration.Key;
                    }
                }
                if (!_namespaces.ContainsKey(worksheet.Prefix)) _namespaces[worksheet.Prefix] = worksheet.NamespaceUri;
            }

            /// <summary>The prefix the root binds to the dyDescent namespace, or null when a row would need a declaration of its own.</summary>
            internal string? DyDescentPrefix { get; }

            internal bool IsMainPrefix(string prefix) =>
                _namespaces.TryGetValue(prefix, out string? uri) && string.Equals(uri, MainNamespace, StringComparison.Ordinal);
        }

        private sealed class GridWriter {
            private readonly ExcelPartBuffer _output;
            private readonly Scope _scope;
            private readonly StringBuilder _text = new StringBuilder(4096);
            private readonly string _row, _cell, _formula, _value;
            private char[] _chars = new char[4096];
            private byte[] _encoded = new byte[4096];

            /// <summary>
            /// Every grid element is written with the grid's own prefix. An element the SDK would write with another
            /// prefix bound to the same namespace is the same XML; reading each element's prefix costs more than writing it.
            /// </summary>
            internal GridWriter(ExcelPartBuffer output, Scope scope, string prefix) {
                _output = output;
                _scope = scope;
                _row = Qualified(prefix, "row");
                _cell = Qualified(prefix, "c");
                _formula = Qualified(prefix, "f");
                _value = Qualified(prefix, "v");
            }

            internal bool TryWriteRows(SheetData sheetData) {
                for (OpenXmlElement? item = sheetData.FirstChild; item != null; item = item.NextSibling()) {
                    if (item is not Row row || !TryWriteRow(row)) return false;
                    Flush();
                }
                return true;
            }

            private bool TryWriteRow(Row row) {
                if (!IsPlain(row) || row.ThickTop != null || row.ThickBot != null || row.ShowPhonetic != null) return false;
                _text.Append('<').Append(_row);
                {
                    if (!TryAppend("r", row.RowIndex) || !TryAppend("spans", row.Spans) || !TryAppend("s", row.StyleIndex)
                        || !TryAppend("customFormat", row.CustomFormat) || !TryAppend("ht", row.Height) || !TryAppend("hidden", row.Hidden)
                        || !TryAppend("customHeight", row.CustomHeight) || !TryAppend("outlineLevel", row.OutlineLevel)
                        || !TryAppend("collapsed", row.Collapsed)) return false;
                    if (row.DyDescent != null) {
                        if (_scope.DyDescentPrefix == null || !TryAppend(_scope.DyDescentPrefix + ":dyDescent", row.DyDescent)) return false;
                    }
                }

                if (row is ExcelStoredRow stored && stored.IsStoreAvailable) return TryWriteStoredRow(stored);

                OpenXmlElement? child = row.FirstChild;
                if (child == null) {
                    _text.Append(" />");
                    return true;
                }
                _text.Append('>');
                for (; child != null; child = child.NextSibling()) {
                    if (child is not Cell cell || !TryWriteCell(cell)) return false;
                }
                _text.Append("</").Append(_row).Append('>');
                return true;
            }

            /// <summary>Writes a row's stored cells merged by column with its DOM cells, as the materialized row would be written.</summary>
            private bool TryWriteStoredRow(ExcelStoredRow row) {
                using (ExcelStoredRow.BeginRawAccess()) {
                    ExcelStoredCell[] cells = row.Cells!;
                    int count = row.StoredCount;
                    string rowText = InvariantNumberText.Get(checked((int)row.RowIndex!.Value));
                    _text.Append('>');
                    int index = 0;
                    for (OpenXmlElement? child = row.FirstChild; child != null; child = child.NextSibling()) {
                        if (child is not Cell cell) return false;
                        int column = ExcelA1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                        if (column <= 0) return false;
                        while (index < count && cells[index].Column < column) {
                            if (!TryWriteStoredCell(cells[index++], rowText)) return false;
                        }
                        if (!TryWriteCell(cell)) return false;
                    }
                    while (index < count) {
                        if (!TryWriteStoredCell(cells[index++], rowText)) return false;
                    }
                    _text.Append("</").Append(_row).Append('>');
                    return true;
                }
            }

            /// <summary>Writes the markup <see cref="TryWriteCell"/> writes for the DOM form of a stored cell, or declines as it would.</summary>
            private bool TryWriteStoredCell(in ExcelStoredCell cell, string rowText) {
                _text.Append('<').Append(_cell).Append(" r=\"");
                AppendColumnLetters(cell.Column);
                _text.Append(rowText).Append('"');
                if (cell.HasStyle) _text.Append(" s=\"").Append(InvariantNumberText.Get(cell.Style)).Append('"');
                if (cell.Type != ExcelStoredCellType.Absent) _text.Append(" t=\"").Append(ExcelStoredRow.TypeName(cell.Type)).Append('"');
                if (!cell.HasValue) {
                    _text.Append(" />");
                    return true;
                }
                _text.Append('>');
                if (!TryWriteText(_value, cell.Value)) return false;
                _text.Append("</").Append(_cell).Append('>');
                return true;
            }

            private void AppendColumnLetters(int column) {
                Span<char> letters = stackalloc char[3];
                int length = 0;
                while (column > 0) {
                    int remainder = (column - 1) % 26;
                    letters[length++] = (char)('A' + remainder);
                    column = (column - 1) / 26;
                }
                while (length > 0) _text.Append(letters[--length]);
            }

            private bool TryWriteCell(Cell cell) {
                if (!IsPlain(cell) || cell.CellMetaIndex != null || cell.ValueMetaIndex != null || cell.ShowPhonetic != null) return false;
                _text.Append('<').Append(_cell);
                if (!TryAppend("r", cell.CellReference) || !TryAppend("s", cell.StyleIndex) || !TryAppend("t", cell.DataType)) return false;

                OpenXmlElement? child = cell.FirstChild;
                if (child == null) {
                    _text.Append(" />");
                    return true;
                }
                _text.Append('>');
                if (child is CellFormula formula) {
                    if (!TryWriteFormula(formula)) return false;
                    child = child.NextSibling();
                }
                if (child is CellValue value) {
                    if (value.MCAttributes != null || value.HasAttributes || !TryWriteText(_value, value.Text)) return false;
                    child = child.NextSibling();
                }
                if (child != null) return false;
                _text.Append("</").Append(_cell).Append('>');
                return true;
            }

            private bool TryWriteFormula(CellFormula formula) {
                if (formula.MCAttributes != null) return false;
                if (!formula.HasAttributes) return TryWriteText(_formula, formula.Text);
                if (formula.ExtendedAttributes.Any()) return false;
                if (formula.AlwaysCalculateArray != null || formula.DataTable2D != null || formula.DataTableRow != null
                    || formula.Input1Deleted != null || formula.Input2Deleted != null || formula.R1 != null || formula.R2 != null
                    || formula.Bx != null || formula.Space != null) return false;
                _text.Append('<').Append(_formula);
                if (!TryAppend("t", formula.FormulaType) || !TryAppend("ref", formula.Reference)
                    || !TryAppend("ca", formula.CalculateCell) || !TryAppend("si", formula.SharedIndex)) return false;
                return TryAppendText(_formula, formula.Text);
            }

            private bool TryWriteText(string name, string? text) {
                _text.Append('<').Append(name);
                return TryAppendText(name, text);
            }

            /// <summary>Closes the start tag and writes the text and end tag. The writer's choice for empty text is left to it.</summary>
            private bool TryAppendText(string name, string? text) {
                if (string.IsNullOrEmpty(text)) return false;
                _text.Append('>');
                if (!TryAppendEscaped(text!, attribute: false)) return false;
                _text.Append("</").Append(name).Append('>');
                return true;
            }

            private bool TryAppend(string name, OpenXmlSimpleType? value) {
                if (value == null) return true;
                string? text = value.InnerText;
                if (text == null) return false;
                _text.Append(' ').Append(name).Append("=\"");
                if (!TryAppendEscaped(text, attribute: true)) return false;
                _text.Append('"');
                return true;
            }

            /// <summary>
            /// The element carries no markup-compatibility or extended attributes. Its own namespace declarations can
            /// be dropped: every name written here uses a prefix the root binds, and grid content holds no qualified names.
            /// </summary>
            private static bool IsPlain(OpenXmlElement element) =>
                element.MCAttributes == null && !element.ExtendedAttributes.Any();

            /// <summary>Escapes as the XML writer does. Characters it would encode or reject are declined instead.</summary>
            private bool TryAppendEscaped(string value, bool attribute) {
                for (int index = 0; index < value.Length; index++) {
                    char character = value[index];
                    if (character >= 0x20 && character < 0x7F) {
                        switch (character) {
                            case '&': _text.Append("&amp;"); break;
                            case '<': _text.Append("&lt;"); break;
                            case '>': _text.Append("&gt;"); break;
                            case '"' when attribute: _text.Append("&quot;"); break;
                            default: _text.Append(character); break;
                        }
                    } else if (character >= 0xA0 && character < 0xD800 || character >= 0xE000 && character < 0xFFFE) {
                        _text.Append(character);
                    } else if (char.IsHighSurrogate(character) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) {
                        _text.Append(character).Append(value[++index]);
                    } else {
                        return false;
                    }
                }
                return true;
            }

            private void Flush() {
                int length = _text.Length;
                if (length == 0) return;
                if (_chars.Length < length) _chars = new char[Math.Max(length, _chars.Length * 2)];
                _text.CopyTo(0, _chars, 0, length);
                int maximum = Utf8.GetMaxByteCount(length);
                if (_encoded.Length < maximum) _encoded = new byte[Math.Max(maximum, _encoded.Length * 2)];
                int written = Utf8.GetBytes(_chars, 0, length, _encoded, 0);
                _output.Write(_encoded, 0, written);
                _text.Clear();
            }
        }
    }
}
