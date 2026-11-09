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

            private bool IsIndexedTag(Utf8Tag tag) =>
                tag.LocalNameStart - tag.NameStart == _worksheetPrefixLength
                && (_worksheetPrefixLength == 0
                    || _buffer!.AsSpan(tag.NameStart, _worksheetPrefixLength).SequenceEqual(
                        _buffer.AsSpan(_worksheetPrefixStart, _worksheetPrefixLength)));
        }
    }
}
