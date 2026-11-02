using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Builds a worksheet's cell grid directly instead of through the SDK's element parser, which costs more than half
    /// of a first edit's time and allocation. The SDK still parses everything around the grid.
    /// </summary>
    /// <remarks>
    /// Only the markup that spreadsheet writers emit for plain rows and cells is accepted. Anything else, including
    /// malformed input, returns false before the part is touched, and the SDK then loads the part and reports its errors.
    /// Attribute text is kept verbatim as the SDK keeps it, so an invalid value still surfaces only when it is read.
    /// </remarks>
    internal static class ExcelWorksheetGridLoader {
        private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string DyDescentNamespaceDeclaration = "xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"";
        private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
        private static readonly string[] CellTypes = { "s", "str", "n", "b", "e" };
        private static readonly string[] FormulaTypes = { "normal", "array", "shared", "dataTable" };

        /// <summary>Loads <paramref name="part"/> and returns true, or leaves it unloaded and returns false.</summary>
        internal static bool TryLoad(WorksheetPart part) {
            var buffer = new ExcelPartBuffer();
            try {
                using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read)) buffer.ReadFrom(stream);
                ReadOnlySpan<byte> xml = new ReadOnlySpan<byte>(buffer.Buffer, 0, buffer.ContentLength);
                if (!TryFindGrid(xml, out int rootStart, out int gridStart, out Tags tags, out bool dyDescentDeclared)) return false;

                var rows = new List<Row>(256);
                var reader = new GridReader(xml, gridStart + tags.GridStart.Length, tags, dyDescentDeclared);
                if (!reader.TryReadRows(rows)) return false;

                string shell = StrictUtf8.GetString(buffer.Buffer, rootStart, gridStart - rootStart) + tags.EmptyGrid
                    + StrictUtf8.GetString(buffer.Buffer, reader.Position, buffer.ContentLength - reader.Position);
                var worksheet = new Worksheet(shell);
                SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData == null) return false;
                foreach (Row row in rows) sheetData.AppendChild(row);
                part.Worksheet = worksheet;
                return true;
            } catch (Exception exception) when (exception is XmlException or ArgumentException or DecoderFallbackException or InvalidOperationException) {
                return false;
            } finally {
                buffer.Release();
            }
        }

        private static bool TryFindGrid(ReadOnlySpan<byte> xml, out int rootStart, out int gridStart, out Tags tags, out bool dyDescentDeclared) {
            rootStart = gridStart = 0;
            tags = null!;
            dyDescentDeclared = false;
            int position = xml.Length >= 3 && xml[0] == 0xEF && xml[1] == 0xBB && xml[2] == 0xBF ? 3 : 0;
            if (GridReader.Matches(xml, position, "<?xml ")) {
                int end = xml.Slice(position).IndexOf((byte)'>');
                if (end < 0) return false;
                string declaration = Encoding.ASCII.GetString(xml.Slice(position, end).ToArray());
                int encoding = declaration.IndexOf("encoding=", StringComparison.Ordinal);
                if (encoding >= 0 && declaration.IndexOf("utf-8", encoding, StringComparison.OrdinalIgnoreCase) != encoding + 10) return false;
                position += end + 1;
            }
            while (position < xml.Length && (xml[position] == (byte)'\r' || xml[position] == (byte)'\n')) position++;

            int rootEnd = xml.Slice(position).IndexOf((byte)'>');
            if (rootEnd < 0) return false;
            string root = Encoding.UTF8.GetString(xml.Slice(position, rootEnd).ToArray());
            if (root.StartsWith("<worksheet", StringComparison.Ordinal)) {
                if (root.IndexOf("xmlns=\"" + MainNamespace + "\"", StringComparison.Ordinal) < 0) return false;
                tags = Tags.Unprefixed;
            } else if (root.StartsWith("<x:worksheet", StringComparison.Ordinal)) {
                if (root.IndexOf("xmlns:x=\"" + MainNamespace + "\"", StringComparison.Ordinal) < 0) return false;
                tags = Tags.XPrefixed;
            } else {
                return false;
            }

            gridStart = xml.Slice(position).IndexOf(tags.GridStart);
            if (gridStart < 0) return false;
            ReadOnlySpan<byte> header = xml.Slice(position, gridStart);
            // A comment, CDATA section or processing instruction could hide or fake the grid's start tag.
            if (header.IndexOf(new[] { (byte)'<', (byte)'!' }) >= 0 || header.IndexOf(new[] { (byte)'<', (byte)'?' }) >= 0) return false;
            dyDescentDeclared = root.IndexOf(DyDescentNamespaceDeclaration, StringComparison.Ordinal) >= 0;
            rootStart = position;
            gridStart += position;
            return true;
        }

        /// <summary>The start and end tags of the grid in the two forms spreadsheet writers use for the main namespace.</summary>
        private sealed class Tags {
            internal static readonly Tags Unprefixed = new Tags(string.Empty);
            internal static readonly Tags XPrefixed = new Tags("x:");

            private Tags(string prefix) {
                GridStart = Encoding.ASCII.GetBytes("<" + prefix + "sheetData>");
                EmptyGrid = "<" + prefix + "sheetData/>";
                GridEnd = "</" + prefix + "sheetData>";
                Row = "<" + prefix + "row";
                RowEnd = "</" + prefix + "row>";
                Cell = "<" + prefix + "c";
                CellEnd = "</" + prefix + "c>";
                Formula = "<" + prefix + "f";
                FormulaEnd = "</" + prefix + "f>";
                Value = "<" + prefix + "v>";
                ValueEnd = "</" + prefix + "v>";
            }

            internal byte[] GridStart { get; }
            internal string EmptyGrid { get; }
            internal string GridEnd { get; }
            internal string Row { get; }
            internal string RowEnd { get; }
            internal string Cell { get; }
            internal string CellEnd { get; }
            internal string Formula { get; }
            internal string FormulaEnd { get; }
            internal string Value { get; }
            internal string ValueEnd { get; }
        }

        private ref struct GridReader {
            private readonly ReadOnlySpan<byte> _xml;
            private readonly Tags _tags;
            private readonly bool _dyDescentDeclared;
            private int _position;

            internal GridReader(ReadOnlySpan<byte> xml, int position, Tags tags, bool dyDescentDeclared) {
                _xml = xml;
                _position = position;
                _tags = tags;
                _dyDescentDeclared = dyDescentDeclared;
            }

            internal int Position => _position;

            internal bool TryReadRows(List<Row> rows) {
                while (true) {
                    if (Skip(_tags.GridEnd)) return true;
                    if (!Skip(_tags.Row)) return false;
                    var row = new Row();
                    if (!TryReadRowAttributes(row, out bool hasCells)) return false;
                    if (hasCells) {
                        while (!Skip(_tags.RowEnd)) {
                            if (!Skip(_tags.Cell) || !TryReadCell(row)) return false;
                        }
                    }
                    rows.Add(row);
                }
            }

            private bool TryReadRowAttributes(Row row, out bool hasContent) {
                while (TryReadAttribute(out ReadOnlySpan<byte> name, out ReadOnlySpan<byte> value, out hasContent, out bool done)) {
                    if (done) return true;
                    string text = Text(value);
                    if (Is(name, "r")) { if (row.RowIndex != null) return false; row.RowIndex = new UInt32Value { InnerText = text }; }
                    else if (Is(name, "spans")) { if (row.Spans != null) return false; row.Spans = new ListValue<StringValue> { InnerText = text }; }
                    else if (Is(name, "s")) { if (row.StyleIndex != null) return false; row.StyleIndex = new UInt32Value { InnerText = text }; }
                    else if (Is(name, "customFormat")) { if (row.CustomFormat != null) return false; row.CustomFormat = new BooleanValue { InnerText = text }; }
                    else if (Is(name, "ht")) { if (row.Height != null) return false; row.Height = new DoubleValue { InnerText = text }; }
                    else if (Is(name, "hidden")) { if (row.Hidden != null) return false; row.Hidden = new BooleanValue { InnerText = text }; }
                    else if (Is(name, "customHeight")) { if (row.CustomHeight != null) return false; row.CustomHeight = new BooleanValue { InnerText = text }; }
                    else if (Is(name, "x14ac:dyDescent")) { if (!_dyDescentDeclared || row.DyDescent != null) return false; row.DyDescent = new DoubleValue { InnerText = text }; }
                    else return false;
                }
                hasContent = false;
                return false;
            }

            private bool TryReadCell(Row row) {
                var cell = new Cell();
                bool hasContent;
                // Duplicates are tracked locally: each SDK attribute getter walks the element's attribute list.
                bool reference = false, style = false, type = false;
                while (true) {
                    if (!TryReadAttribute(out ReadOnlySpan<byte> name, out ReadOnlySpan<byte> value, out hasContent, out bool done)) return false;
                    if (done) break;
                    if (Is(name, "r")) {
                        if (reference) return false;
                        reference = true;
                        cell.CellReference = Text(value);
                    } else if (Is(name, "s")) {
                        if (style) return false;
                        style = true;
                        cell.StyleIndex = new UInt32Value { InnerText = Text(value) };
                    } else if (Is(name, "t")) {
                        if (type || !IsAnyOf(value, CellTypes)) return false;
                        type = true;
                        cell.DataType = new EnumValue<CellValues> { InnerText = Text(value) };
                    } else {
                        return false;
                    }
                }

                if (hasContent) {
                    if (Skip(_tags.Formula) && !TryReadFormula(cell)) return false;
                    if (Skip(_tags.Value)) {
                        if (!TryReadText(_tags.ValueEnd, out string? text)) return false;
                        cell.AppendChild(new CellValue(text!));
                    }
                    if (!Skip(_tags.CellEnd)) return false;
                }
                row.AppendChild(cell);
                return true;
            }

            private bool TryReadFormula(Cell cell) {
                var formula = new CellFormula();
                bool hasContent;
                bool type = false, reference = false, sharedIndex = false, calculate = false;
                while (true) {
                    if (!TryReadAttribute(out ReadOnlySpan<byte> name, out ReadOnlySpan<byte> value, out hasContent, out bool done)) return false;
                    if (done) break;
                    if (Is(name, "t")) {
                        if (type || !IsAnyOf(value, FormulaTypes)) return false;
                        type = true;
                        formula.FormulaType = new EnumValue<CellFormulaValues> { InnerText = Text(value) };
                    } else if (Is(name, "ref")) {
                        if (reference) return false;
                        reference = true;
                        formula.Reference = Text(value);
                    } else if (Is(name, "si")) {
                        if (sharedIndex) return false;
                        sharedIndex = true;
                        formula.SharedIndex = new UInt32Value { InnerText = Text(value) };
                    } else if (Is(name, "ca")) {
                        if (calculate) return false;
                        calculate = true;
                        formula.CalculateCell = new BooleanValue { InnerText = Text(value) };
                    } else {
                        return false;
                    }
                }

                if (hasContent) {
                    if (!TryReadText(_tags.FormulaEnd, out string? text)) return false;
                    formula.Text = text!;
                }
                cell.AppendChild(formula);
                return true;
            }

            /// <summary>Reads one attribute, or the end of the start tag, which sets <paramref name="done"/>.</summary>
            private bool TryReadAttribute(out ReadOnlySpan<byte> name, out ReadOnlySpan<byte> value, out bool hasContent, out bool done) {
                name = value = default;
                hasContent = done = false;
                if (_position >= _xml.Length) return false;
                byte next = _xml[_position];
                if (next == (byte)'>') { _position++; hasContent = done = true; return true; }
                if (next == (byte)'/') { done = true; return Skip("/>"); }
                if (next != (byte)' ') return false;
                _position++;

                int equals = _xml.Slice(_position).IndexOf((byte)'=');
                if (equals <= 0 || _position + equals + 1 >= _xml.Length || _xml[_position + equals + 1] != (byte)'"') return false;
                name = _xml.Slice(_position, equals);
                int start = _position + equals + 2;
                int length = _xml.Slice(start).IndexOf((byte)'"');
                if (length < 0) return false;
                value = _xml.Slice(start, length);
                foreach (byte character in value) {
                    // Escapes and whitespace normalization stay with the XML reader, as does anything outside ASCII.
                    if (character < 0x20 || character >= 0x7F || character == (byte)'<' || character == (byte)'&') return false;
                }
                _position = start + length + 1;
                return true;
            }

            private bool TryReadText(string endTag, out string? text) {
                text = null;
                int length = _xml.Slice(_position).IndexOf((byte)'<');
                if (length <= 0) return false;
                ReadOnlySpan<byte> raw = _xml.Slice(_position, length);
                bool escaped = false;
                foreach (byte character in raw) {
                    // A raw '>' could close "]]>", and a carriage return would need XML line-end normalization.
                    if ((character < 0x20 && character != (byte)'\t' && character != (byte)'\n') || character >= 0x7F || character == (byte)'>') return false;
                    if (character == (byte)'&') escaped = true;
                }
                if (raw[0] <= (byte)' ' || raw[raw.Length - 1] <= (byte)' ') return false;
                _position += length;
                if (!Skip(endTag)) return false;
                if (!escaped) { text = Text(raw); return true; }
                text = TryDecode(raw);
                return text != null;
            }

            private static string? TryDecode(ReadOnlySpan<byte> raw) {
                var text = new StringBuilder(raw.Length);
                for (int index = 0; index < raw.Length; index++) {
                    if (raw[index] != (byte)'&') { text.Append((char)raw[index]); continue; }
                    ReadOnlySpan<byte> rest = raw.Slice(index);
                    if (Matches(rest, 0, "&quot;")) { text.Append('"'); index += 5; }
                    else if (Matches(rest, 0, "&amp;")) { text.Append('&'); index += 4; }
                    else if (Matches(rest, 0, "&lt;")) { text.Append('<'); index += 3; }
                    else if (Matches(rest, 0, "&gt;")) { text.Append('>'); index += 3; }
                    else if (Matches(rest, 0, "&apos;")) { text.Append('\''); index += 5; }
                    else return null;
                }
                return text.ToString();
            }

            private bool Skip(string token) {
                if (!Matches(_xml, _position, token)) return false;
                _position += token.Length;
                return true;
            }

            internal static bool Matches(ReadOnlySpan<byte> xml, int position, string token) {
                if (xml.Length - position < token.Length) return false;
                for (int index = 0; index < token.Length; index++) {
                    if (xml[position + index] != token[index]) return false;
                }
                return true;
            }

            private static bool Is(ReadOnlySpan<byte> name, string expected) => name.Length == expected.Length && Matches(name, 0, expected);

            private static bool IsAnyOf(ReadOnlySpan<byte> value, string[] expected) {
                foreach (string candidate in expected) {
                    if (Is(value, candidate)) return true;
                }
                return false;
            }

            private static string Text(ReadOnlySpan<byte> ascii) {
    #if NET5_0_OR_GREATER
                return Encoding.ASCII.GetString(ascii);
    #else
                var characters = new char[ascii.Length];
                for (int index = 0; index < characters.Length; index++) characters[index] = (char)ascii[index];
                return new string(characters);
    #endif
            }
        }
    }
}
