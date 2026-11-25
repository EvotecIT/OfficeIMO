using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>Applies insertions and replacements to a part's bytes, leaving everything else as it was.</summary>
    internal sealed class ExcelPartEdits {
        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        private readonly List<(int Position, int RemoveLength, byte[] Insert)> _edits = new();

        internal int Count => _edits.Count;

        internal void Replace(int position, int removeLength, string text) => _edits.Add((position, removeLength, Utf8.GetBytes(text)));

        internal void Insert(int position, string text) => Replace(position, 0, text);

        internal byte[] Apply(byte[] source, int length, out int resultLength) {
            _edits.Sort(static (left, right) => left.Position.CompareTo(right.Position));
            int size = length;
            foreach ((int _, int remove, byte[] insert) in _edits) size += insert.Length - remove;
            var output = new byte[size];
            int copied = 0, written = 0;
            foreach ((int position, int remove, byte[] insert) in _edits) {
                if (position < copied) throw new InvalidOperationException("Part edits overlap.");
                Buffer.BlockCopy(source, copied, output, written, position - copied);
                written += position - copied;
                Buffer.BlockCopy(insert, 0, output, written, insert.Length);
                written += insert.Length;
                copied = position + remove;
            }
            Buffer.BlockCopy(source, copied, output, written, length - copied);
            resultLength = written + length - copied;
            return output;
        }

        /// <summary>Sets an unqualified attribute on the start tag the reader is on, keeping its position when it exists.</summary>
        internal void SetAttribute(OfficeXmlPullReader reader, string localName, string value) {
            for (int index = 0; index < reader.AttributeCount; index++) {
                if (reader.GetAttributeNamespace(index, out _) == OfficeXmlNamespaces.None
                    && string.Equals(reader.GetAttributeLocalName(index), localName, StringComparison.Ordinal)) {
                    reader.GetAttributeRawValue(index, out int start, out int length);
                    Replace(start, length, value);
                    return;
                }
            }
            Insert(StartTagNameEnd(reader), " " + localName + "=\"" + value + "\"");
        }

        // New attributes follow the element name, so they never land inside an empty element's "/>".
        private static int StartTagNameEnd(OfficeXmlPullReader reader) {
            byte[] data = reader.Data;
            int position = reader.MarkupStart + 1;
            while (position < reader.MarkupEnd && data[position] != (byte)' ' && data[position] != (byte)'\t' && data[position] != (byte)'\r'
                && data[position] != (byte)'\n' && data[position] != (byte)'>' && data[position] != (byte)'/') {
                position++;
            }
            return position;
        }

        internal static string Escape(string value) {
            var text = new StringBuilder(value.Length + 8);
            foreach (char character in value) {
                switch (character) {
                    case '&': text.Append("&amp;"); break;
                    case '<': text.Append("&lt;"); break;
                    case '>': text.Append("&gt;"); break;
                    case '"': text.Append("&quot;"); break;
                    default: text.Append(character); break;
                }
            }
            return text.ToString();
        }
    }

    /// <summary>
    /// The shared-string table as the model's save keeps it: existing items found by their text (a later duplicate
    /// wins, as the model's cache keeps it), new items appended in the order values were written.
    /// </summary>
    internal sealed class ExcelDeferredSharedStrings {
        private readonly Dictionary<string, int> _indexes = new(StringComparer.Ordinal);
        private readonly List<string> _added = new();
        private int _count;

        private ExcelDeferredSharedStrings() {
        }

        /// <summary>The part, or null when the workbook has none yet.</summary>
        internal string? PartName { get; private set; }

        internal int AddedReferences { get; set; }

        internal int RemovedReferences { get; set; }

        internal bool Changed => _added.Count > 0 || AddedReferences != 0 || RemovedReferences != 0;

        /// <summary>Reads the table, or returns null when its markup needs the model.</summary>
        internal static ExcelDeferredSharedStrings? Load(OfficeOpcEditablePackage package, string? partName) {
            var table = new ExcelDeferredSharedStrings { PartName = partName };
            if (partName == null) return table;
            OfficePooledBytes content = package.ReadPart(partName);
            using (content) {
                var reader = new OfficeXmlPullReader(content.Array, content.Length, processMarkupCompatibility: false);
                if (reader.Transcoded) return null;
                string? text = null;
                StringBuilder? runs = null;
                bool inItem = false, inRun = false, root = false;
                while (reader.Read()) {
                    int depth = reader.Depth;
                    if (reader.NodeKind == OfficeXmlNodeKind.StartElement) {
                        if (depth == 1) {
                            if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "sst")
                                || reader.ElementUriId != OfficeXmlNamespaces.TransitionalUriId(OfficeXmlNamespaces.SpreadsheetML)) {
                                return null;
                            }
                            root = true;
                        } else if (depth == 2 && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "si")) {
                            inItem = true;
                            text = null;
                            runs = null;
                        } else if (inItem && depth == 3 && reader.NamespaceId == OfficeXmlNamespaces.SpreadsheetML) {
                            if (reader.LocalNameEquals("t")) {
                                // The item's own text wins over its runs, as the model reads it.
                                text ??= ReadElementText(reader);
                            } else if (reader.LocalNameEquals("r")) {
                                inRun = true;
                            }
                        } else if (inRun && depth == 4 && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "t")) {
                            (runs ??= new StringBuilder()).Append(Utilities.ExcelSanitizer.DecodeXString(ReadElementText(reader)));
                        }
                    } else if (reader.NodeKind == OfficeXmlNodeKind.EndElement) {
                        if (depth == 3) inRun = false;
                        else if (depth == 2 && inItem) {
                            inItem = false;
                            string value = text != null ? Utilities.ExcelSanitizer.DecodeXString(text) : runs?.ToString() ?? string.Empty;
                            table._indexes[value] = table._count++;
                        }
                    }
                }
                return root ? table : null;
            }
        }

        private static string ReadElementText(OfficeXmlPullReader reader) {
            if (reader.IsEmptyElement) return string.Empty;
            var text = new StringBuilder();
            int depth = reader.Depth;
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) break;
                if (reader.NodeKind == OfficeXmlNodeKind.Text && reader.Depth == depth + 1) text.Append(reader.GetText());
                // The model's text stops at its first child element.
                if (reader.NodeKind == OfficeXmlNodeKind.StartElement) return text.ToString();
            }
            return text.ToString();
        }

        /// <summary>Finds or adds the item for <paramref name="text"/>; false past <paramref name="addLimit"/> items.</summary>
        internal bool TryGetOrAdd(string text, int addLimit, out int index) {
            if (_indexes.TryGetValue(text, out index)) return true;
            if (_count >= addLimit) {
                index = -1;
                return false;
            }
            index = _count++;
            _indexes[text] = index;
            _added.Add(text);
            return true;
        }

        /// <summary>Writes the table with its new items, or returns false when the markup needs the model.</summary>
        internal bool TryWrite(OfficeOpcEditablePackage package, string workbookPartName) {
            if (!Changed) return true;
            if (PartName == null) {
                // The model creates the table part at the usual name, which must then be free.
                const string name = "/xl/sharedStrings.xml";
                if (package.PartExists(name)) return false;
                var xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"utf-8\"?><x:sst xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
                xml.Append(" count=\"").Append(Count(0, 0).ToString(CultureInfo.InvariantCulture)).Append("\" uniqueCount=\"")
                    .Append(_count.ToString(CultureInfo.InvariantCulture)).Append("\">");
                foreach (string item in _added) AppendItem(xml, "x:", item);
                xml.Append("</x:sst>");
                package.CreateBinaryPart(name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                    Encoding.UTF8.GetBytes(xml.ToString()));
                package.AddRelationship(workbookPartName, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "sharedStrings",
                    RelativeTarget(workbookPartName, name));
                return true;
            }

            OfficePooledBytes content = package.ReadPart(PartName);
            byte[] result;
            int resultLength;
            using (content) {
                var reader = new OfficeXmlPullReader(content.Array, content.Length, processMarkupCompatibility: false);
                var edits = new ExcelPartEdits();
                string prefix = string.Empty;
                int declaredCount = -1;
                while (reader.Read()) {
                    if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == 1) {
                        prefix = reader.ElementPrefix.Length == 0 ? string.Empty : reader.ElementPrefix + ":";
                        if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "count", out string countText)
                            && int.TryParse(countText, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed)) {
                            declaredCount = parsed;
                        }
                        int originalItems = _count - _added.Count;
                        edits.SetAttribute(reader, "count", Count(declaredCount < 0 ? originalItems : declaredCount, originalItems).ToString(CultureInfo.InvariantCulture));
                        edits.SetAttribute(reader, "uniqueCount", _count.ToString(CultureInfo.InvariantCulture));
                        if (reader.IsEmptyElement) {
                            if (_added.Count == 0) continue;
                            // <sst/> becomes an element holding the new items.
                            int slash = reader.MarkupEnd - 2;
                            while (slash > reader.MarkupStart && content.Array[slash] != (byte)'/') slash--;
                            var items = new StringBuilder(">");
                            foreach (string item in _added) AppendItem(items, prefix, item);
                            items.Append("</").Append(prefix).Append("sst");
                            edits.Replace(slash, 1, items.ToString());
                        }
                        continue;
                    }
                    if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 1 && _added.Count > 0
                        && reader.MarkupEnd > reader.MarkupStart) {
                        var items = new StringBuilder();
                        foreach (string item in _added) AppendItem(items, prefix, item);
                        edits.Insert(reader.MarkupStart, items.ToString());
                    }
                }
                result = edits.Apply(content.Array, content.Length, out resultLength);
            }
            package.ReplacePart(PartName, result, resultLength);
            return true;
        }

        private static string RelativeTarget(string sourcePartName, string targetPartName) {
            int slash = sourcePartName.LastIndexOf('/');
            string directory = sourcePartName.Substring(0, slash + 1);
            return targetPartName.StartsWith(directory, StringComparison.Ordinal) ? targetPartName.Substring(directory.Length) : targetPartName;
        }

        // The model counts every shared-string cell; this keeps the declared count up to date with the edits.
        private int Count(int declared, int originalItems) => Math.Max(_count, declared + AddedReferences - RemovedReferences);

        private static void AppendItem(StringBuilder xml, string prefix, string text) {
            string encoded = Utilities.ExcelSanitizer.EncodeXString(text);
            bool preserve = text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]));
            xml.Append('<').Append(prefix).Append("si><").Append(prefix).Append('t');
            if (preserve) xml.Append(" xml:space=\"preserve\"");
            xml.Append('>').Append(ExcelPartEdits.Escape(encoded)).Append("</").Append(prefix).Append("t></").Append(prefix).Append("si>");
        }
    }

    /// <summary>
    /// The automatic date format the model gives a date value: the cell's own style when it already shows dates,
    /// otherwise its style with the built-in short date (14), reused when the stylesheet already has that format.
    /// </summary>
    internal sealed class ExcelDeferredDateStyles {
        // xf attributes in schema order, which is how the model writes and compares a format.
        private static readonly string[] FormatAttributeOrder = {
            "numFmtId", "fontId", "fillId", "borderId", "xfId", "quotePrefix", "pivotButton", "applyNumberFormat", "applyFont",
            "applyFill", "applyBorder", "applyAlignment", "applyProtection"
        };

        private readonly OfficeOpcEditablePackage _package;
        private readonly string _partName;
        private readonly List<Format> _formats = new();
        private readonly Dictionary<string, uint> _formatIndexes = new(StringComparer.Ordinal);
        private readonly Dictionary<uint, string> _numberFormatCodes = new();
        private readonly Dictionary<uint, uint> _derived = new();
        private readonly List<Format> _added = new();
        private readonly bool _primitivesComplete;
        private int _cellFormatsEnd = -1;
        private string _prefix = string.Empty;

        private ExcelDeferredDateStyles(OfficeOpcEditablePackage package, string partName, bool primitivesComplete) {
            _package = package;
            _partName = partName;
            _primitivesComplete = primitivesComplete;
        }

        internal bool Changed => _added.Count > 0;

        private sealed class Format {
            internal List<(string Name, string Value)> Attributes = new();
            internal string Children = string.Empty;
            internal string Key = string.Empty;
        }

        /// <summary>Reads the stylesheet, or returns null when it needs the model.</summary>
        internal static ExcelDeferredDateStyles? Load(OfficeOpcEditablePackage package, string? partName) {
            if (partName == null) return null;
            OfficePooledBytes content = package.ReadPart(partName);
            using (content) {
                var reader = new OfficeXmlPullReader(content.Array, content.Length, processMarkupCompatibility: false);
                if (reader.Transcoded) return null;
                var counts = new Dictionary<string, (int Declared, int Actual)>(StringComparer.Ordinal);
                bool firstFontComplete = false, hasNoneFill = false, hasGrayFill = false;
                int fontIndex = -1;
                bool firstFontHasSize = false, firstFontHasName = false;
                string? section = null;
                ExcelDeferredDateStyles? styles = null;
                while (reader.Read()) {
                    int depth = reader.Depth;
                    if (reader.NodeKind == OfficeXmlNodeKind.EndElement) {
                        if (depth == 2 && section == "cellXfs" && styles != null) styles._cellFormatsEnd = reader.MarkupStart;
                        if (depth == 2) section = null;
                        if (depth == 3 && section == "fonts" && fontIndex == 0) firstFontComplete = firstFontHasSize && firstFontHasName;
                        continue;
                    }
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                    if (depth == 1) {
                        if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "styleSheet")
                            || reader.ElementUriId != OfficeXmlNamespaces.TransitionalUriId(OfficeXmlNamespaces.SpreadsheetML)) {
                            return null;
                        }
                        styles = new ExcelDeferredDateStyles(package, partName, primitivesComplete: false) {
                            _prefix = reader.ElementPrefix.Length == 0 ? string.Empty : reader.ElementPrefix + ":"
                        };
                        continue;
                    }
                    if (styles == null) return null;
                    if (depth == 2) {
                        section = reader.NamespaceId == OfficeXmlNamespaces.SpreadsheetML ? reader.LocalName : null;
                        if (section != null) {
                            int declared = reader.TryGetAttribute(OfficeXmlNamespaces.None, "count", out string countText)
                                && int.TryParse(countText, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) ? parsed : -1;
                            if (counts.ContainsKey(section)) return null;
                            counts[section] = (declared, 0);
                        }
                        if (section == "cellXfs" && reader.IsEmptyElement) return null;
                        continue;
                    }
                    if (depth == 3 && section != null && counts.TryGetValue(section, out var count)) {
                        string child = reader.NamespaceId == OfficeXmlNamespaces.SpreadsheetML ? reader.LocalName : string.Empty;
                        bool counted = section switch {
                            "fonts" => child == "font",
                            "fills" => child == "fill",
                            "borders" => child == "border",
                            "cellStyleXfs" or "cellXfs" => child == "xf",
                            "cellStyles" => child == "cellStyle",
                            "dxfs" => child == "dxf",
                            "tableStyles" => child == "tableStyle",
                            "numFmts" => child == "numFmt",
                            _ => false
                        };
                        if (counted) counts[section] = (count.Declared, count.Actual + 1);
                        if (section == "fonts" && counted) fontIndex++;
                        if (section == "numFmts" && counted
                            && reader.TryGetAttribute(OfficeXmlNamespaces.None, "numFmtId", out string idText)
                            && uint.TryParse(idText, NumberStyles.None, CultureInfo.InvariantCulture, out uint id)
                            && !styles._numberFormatCodes.ContainsKey(id)) {
                            styles._numberFormatCodes[id] = reader.TryGetAttribute(OfficeXmlNamespaces.None, "formatCode", out string code) ? code : string.Empty;
                        }
                        if (section == "cellXfs" && counted) styles.ReadFormat(reader);
                        continue;
                    }
                    if (depth == 4 && section == "fonts" && fontIndex == 0 && reader.NamespaceId == OfficeXmlNamespaces.SpreadsheetML) {
                        if (reader.LocalNameEquals("sz")) firstFontHasSize = true;
                        else if (reader.LocalNameEquals("name")) firstFontHasName = true;
                    }
                    if (depth == 4 && section == "fills" && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "patternFill")
                        && reader.TryGetAttribute(OfficeXmlNamespaces.None, "patternType", out string pattern)) {
                        hasNoneFill |= pattern == "none";
                        hasGrayFill |= pattern == "gray125";
                    }
                }
                if (styles == null || styles._cellFormatsEnd < 0) return null;
                // The model completes a stylesheet before it adds a format; one that needs completing stays with the model.
                bool complete = firstFontComplete && hasNoneFill && hasGrayFill;
                foreach (string required in new[] { "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles" }) {
                    if (!counts.TryGetValue(required, out var count) || count.Declared != count.Actual
                        || count.Actual == 0 && required is not ("dxfs" or "tableStyles")) {
                        complete = false;
                    }
                }
                if (counts.TryGetValue("numFmts", out var numberFormats) && numberFormats.Declared != numberFormats.Actual) complete = false;
                var loaded = new ExcelDeferredDateStyles(package, partName, complete) {
                    _prefix = styles._prefix,
                    _cellFormatsEnd = styles._cellFormatsEnd
                };
                loaded._formats.AddRange(styles._formats);
                foreach (var pair in styles._formatIndexes) loaded._formatIndexes[pair.Key] = pair.Value;
                foreach (var pair in styles._numberFormatCodes) loaded._numberFormatCodes[pair.Key] = pair.Value;
                return loaded;
            }
        }

        private void ReadFormat(OfficeXmlPullReader reader) {
            var format = new Format();
            for (int index = 0; index < reader.AttributeCount; index++) {
                int namespaceId = reader.GetAttributeNamespace(index, out _);
                if (namespaceId == OfficeXmlNamespaces.Xmlns) continue;
                string name = reader.GetAttributeLocalName(index);
                if (namespaceId != OfficeXmlNamespaces.None) name = "{" + namespaceId.ToString(CultureInfo.InvariantCulture) + "}" + name;
                format.Attributes.Add((name, reader.GetAttributeValue(index)));
            }
            if (!reader.IsEmptyElement) {
                int childrenStart = reader.MarkupEnd;
                int depth = reader.Depth;
                while (reader.Read()) {
                    if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) break;
                }
                format.Children = Encoding.UTF8.GetString(reader.Data, childrenStart, reader.MarkupStart - childrenStart);
            } else {
                reader.Read();
            }
            format.Key = KeyOf(format);
            uint index2 = (uint)_formats.Count;
            _formats.Add(format);
            if (!_formatIndexes.ContainsKey(format.Key)) _formatIndexes[format.Key] = index2;
        }

        private static string KeyOf(Format format) {
            var key = new StringBuilder();
            foreach (string name in FormatAttributeOrder) {
                foreach ((string attributeName, string value) in format.Attributes) {
                    if (attributeName == name) key.Append(name).Append('=').Append(value).Append(';');
                }
            }
            foreach ((string attributeName, string value) in format.Attributes) {
                if (Array.IndexOf(FormatAttributeOrder, attributeName) < 0) key.Append(attributeName).Append('=').Append(value).Append(';');
            }
            return key.Append('|').Append(format.Children).ToString();
        }

        /// <summary>The style a date takes in a cell whose base style is <paramref name="baseStyle"/>, or null when the model must add it.</summary>
        internal uint? DateStyle(uint baseStyle) {
            if (_derived.TryGetValue(baseStyle, out uint known)) return known;
            uint? result = baseStyle != 0U && HasCalendarFormat(baseStyle) ? baseStyle : Derive(baseStyle);
            if (result.HasValue) _derived[baseStyle] = result.Value;
            return result;
        }

        private bool HasCalendarFormat(uint styleIndex) {
            if (styleIndex >= (uint)_formats.Count) return false;
            string? numberFormatText = Attribute(_formats[(int)styleIndex], "numFmtId");
            if (numberFormatText == null || !uint.TryParse(numberFormatText, NumberStyles.None, CultureInfo.InvariantCulture, out uint numberFormatId)
                || numberFormatId == 0U) {
                return false;
            }
            string? code = _numberFormatCodes.TryGetValue(numberFormatId, out string? custom) ? custom : ExcelBuiltInNumberFormats.GetCode(numberFormatId);
            return ExcelNumberFormatClassifier.ClassifySerialDateStyle(numberFormatId, code) == ExcelSerialDateStyle.Calendar;
        }

        private uint? Derive(uint baseStyle) {
            var candidate = new Format();
            if (baseStyle < (uint)_formats.Count) {
                Format source = _formats[(int)baseStyle];
                candidate.Attributes.AddRange(source.Attributes);
                candidate.Children = source.Children;
            } else {
                candidate.Attributes.Add(("numFmtId", "0"));
                candidate.Attributes.Add(("fontId", "0"));
                candidate.Attributes.Add(("fillId", "0"));
                candidate.Attributes.Add(("borderId", "0"));
                candidate.Attributes.Add(("xfId", "0"));
            }
            Set(candidate, "numFmtId", "14");
            Set(candidate, "applyNumberFormat", "1");
            candidate.Key = KeyOf(candidate);
            if (_formatIndexes.TryGetValue(candidate.Key, out uint existing)) return existing;
            // A new format means the model first completes the stylesheet; one that would change stays with it.
            if (!_primitivesComplete || candidate.Children.Length > 0 || candidate.Attributes.Exists(static attribute => attribute.Name.StartsWith("{", StringComparison.Ordinal))) {
                return null;
            }
            uint index = (uint)_formats.Count;
            _formats.Add(candidate);
            _added.Add(candidate);
            _formatIndexes[candidate.Key] = index;
            return index;
        }

        private static string? Attribute(Format format, string name) {
            foreach ((string attributeName, string value) in format.Attributes) {
                if (attributeName == name) return value;
            }
            return null;
        }

        private static void Set(Format format, string name, string value) {
            int index = format.Attributes.FindIndex(attribute => attribute.Name == name);
            if (index >= 0) format.Attributes[index] = (name, value);
            else format.Attributes.Add((name, value));
        }

        /// <summary>Writes the stylesheet with its new formats.</summary>
        internal void Write() {
            if (_added.Count == 0) return;
            OfficePooledBytes content = _package.ReadPart(_partName);
            byte[] result;
            int resultLength;
            using (content) {
                var reader = new OfficeXmlPullReader(content.Array, content.Length, processMarkupCompatibility: false);
                var edits = new ExcelPartEdits();
                while (reader.Read()) {
                    if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == 2
                        && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "cellXfs")) {
                        edits.SetAttribute(reader, "count", _formats.Count.ToString(CultureInfo.InvariantCulture));
                        break;
                    }
                }
                var xml = new StringBuilder();
                foreach (Format format in _added) {
                    xml.Append('<').Append(_prefix).Append("xf");
                    foreach (string name in FormatAttributeOrder) {
                        if (Attribute(format, name) is { } value) xml.Append(' ').Append(name).Append("=\"").Append(ExcelPartEdits.Escape(value)).Append('"');
                    }
                    foreach ((string name, string value) in format.Attributes) {
                        if (Array.IndexOf(FormatAttributeOrder, name) < 0) xml.Append(' ').Append(name).Append("=\"").Append(ExcelPartEdits.Escape(value)).Append('"');
                    }
                    xml.Append(" />");
                }
                edits.Insert(_cellFormatsEnd, xml.ToString());
                result = edits.Apply(content.Array, content.Length, out resultLength);
            }
            _package.ReplacePart(_partName, result, resultLength);
        }
    }
}
