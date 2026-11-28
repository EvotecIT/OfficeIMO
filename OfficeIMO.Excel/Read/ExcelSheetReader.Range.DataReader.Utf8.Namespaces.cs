namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private int _worksheetPrefixStart;
            private int _worksheetPrefixLength;

            // Index only one consistent element prefix. Publication still requires
            // complete character, XML structure and resolved namespace proof.
            private void SetWorksheetPrefix(Utf8Tag root) {
                _worksheetPrefixStart = root.NameStart;
                _worksheetPrefixLength = root.LocalNameStart - root.NameStart;
            }

            /// <summary>
            /// Accepts a cell child that binds its own prefix to SpreadsheetML, as the Open XML SDK writes an element
            /// created in code under a root with a default namespace (<c>&lt;x:f xmlns:x="..."&gt;</c>). Only the general
            /// cell scan uses it, so the worksheet still receives the complete XmlReader and namespace proof.
            /// </summary>
            private bool IsSelfBoundSpreadsheetTag(Utf8Tag tag) {
                int prefixLength = tag.LocalNameStart - tag.NameStart - 1;
                if (tag.IsEnd || prefixLength <= 0 || IsIndexedTag(tag)) return false;
                string declaration = "xmlns:" + System.Text.Encoding.ASCII.GetString(_buffer!, tag.NameStart, prefixLength);
                return TryGetAttribute(tag, declaration, out bool found, out int start, out int length) && found
                    && (AsciiEquals(start, length, SpreadsheetNamespace) || AsciiEquals(start, length, StrictSpreadsheetNamespace));
            }

            /// <summary>
            /// Whether the worksheet root binds the prefix its rows and cells use to SpreadsheetML, and <c>sheetData</c>
            /// keeps that binding. The fast checks match element names only, so rows are returned before the full check
            /// only under this binding.
            /// </summary>
            private bool WorksheetRootBindsSpreadsheetNamespace() {
                // The batches start before the first window moves, so the buffer still begins with the worksheet.
                int position = 0;
                if (_buffer == null || !TryReadNextTag(ref position, _length, out Utf8Tag root) || root.IsEnd) return false;
                int prefixLength = root.LocalNameStart - root.NameStart;
                string declaration = prefixLength == 0
                    ? "xmlns"
                    : "xmlns:" + System.Text.Encoding.ASCII.GetString(_buffer, root.NameStart, prefixLength - 1);
                if (!BindsSpreadsheetNamespace(root, declaration, required: true)) return false;
                while (TryReadNextTag(ref position, _length, out Utf8Tag tag)) {
                    if (tag.IsEnd || !LocalNameEquals(tag, "sheetData")) continue;
                    return BindsSpreadsheetNamespace(tag, declaration, required: false);
                }
                return false;
            }

            private bool BindsSpreadsheetNamespace(Utf8Tag tag, string declaration, bool required) {
                if (!TryGetAttribute(tag, declaration, out bool found, out int start, out int length)) return false;
                if (!found) return !required;
                return AsciiEquals(start, length, SpreadsheetNamespace) || AsciiEquals(start, length, StrictSpreadsheetNamespace);
            }

            private bool IsIndexedTag(Utf8Tag tag) =>
                tag.LocalNameStart - tag.NameStart == _worksheetPrefixLength
                && (_worksheetPrefixLength == 0
                    || _buffer!.AsSpan(tag.NameStart, _worksheetPrefixLength).SequenceEqual(
                        _buffer.AsSpan(_worksheetPrefixStart, _worksheetPrefixLength)));
        }
    }
}
