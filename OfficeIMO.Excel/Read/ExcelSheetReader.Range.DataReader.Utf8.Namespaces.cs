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

            private bool IsIndexedTag(Utf8Tag tag) =>
                tag.LocalNameStart - tag.NameStart == _worksheetPrefixLength
                && (_worksheetPrefixLength == 0
                    || _buffer!.AsSpan(tag.NameStart, _worksheetPrefixLength).SequenceEqual(
                        _buffer.AsSpan(_worksheetPrefixStart, _worksheetPrefixLength)));
        }
    }
}
