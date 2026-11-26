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
            // Edits at the same position keep the order they were made in.
            var ordered = _edits.Select((edit, index) => (edit, index)).OrderBy(static item => item.edit.Position).ThenBy(static item => item.index)
                .Select(static item => item.edit).ToList();
            int size = length;
            foreach ((int _, int remove, byte[] insert) in ordered) size += insert.Length - remove;
            var output = new byte[size];
            int copied = 0, written = 0;
            foreach ((int position, int remove, byte[] insert) in ordered) {
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
                    // Written as references so attribute value normalization keeps them as read.
                    case '\t': text.Append("&#x9;"); break;
                    case '\r': text.Append("&#xD;"); break;
                    case '\n': text.Append("&#xA;"); break;
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
                        } else if (depth == 2 && !reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "extLst")) {
                            // Compatibility markup can supply items of its own, which would move the indexes after it.
                            return null;
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
                // Text reports the depth of the element that contains it.
                if (reader.NodeKind == OfficeXmlNodeKind.Text && reader.Depth == depth) text.Append(reader.GetText());
                // The model's text stops at its first child element.
                if (reader.NodeKind == OfficeXmlNodeKind.StartElement) return text.ToString();
            }
            return text.ToString();
        }

        /// <summary>The index of the item for <paramref name="text"/>, or -1 when the table has none.</summary>
        internal int IndexOf(string text) => _indexes.TryGetValue(text, out int index) ? index : -1;

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
                bool inserted = false;
                while (reader.Read()) {
                    if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == 1) {
                        prefix = reader.ElementPrefix.Length == 0 ? string.Empty : reader.ElementPrefix + ":";
                        if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "count", out string countText)
                            && int.TryParse(countText, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed)) {
                            declaredCount = parsed;
                        }
                        int originalItems = _count - _added.Count;
                        // Without a declared reference count the workbook's total is unknown, so it stays undeclared.
                        if (declaredCount >= 0) edits.SetAttribute(reader, "count", Count(declaredCount, originalItems).ToString(CultureInfo.InvariantCulture));
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
                    // New items go after the last item: before the extension list, which closes the table, or before </sst>.
                    bool extensions = reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == 2
                        && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "extLst");
                    bool end = reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 1 && reader.MarkupEnd > reader.MarkupStart;
                    if ((extensions || end) && !inserted && _added.Count > 0) {
                        inserted = true;
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
}
