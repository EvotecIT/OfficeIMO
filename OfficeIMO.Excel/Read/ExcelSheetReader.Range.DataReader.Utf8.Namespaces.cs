#nullable enable

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private int _worksheetPrefixStart;
            private int _worksheetPrefixLength;

            // Index only one consistent element prefix. XmlReader still validates the
            // entire document and the resolved namespaces before any indexed row is exposed.
            private void SetWorksheetPrefix(Utf8Tag root) {
                _worksheetPrefixStart = root.NameStart;
                _worksheetPrefixLength = root.LocalNameStart - root.NameStart;
                if (_worksheetPrefixLength != 0) {
                    _sheetDataSupportsFastValidation = false;
                }
            }

            private bool IsIndexedTag(Utf8Tag tag) =>
                tag.LocalNameStart - tag.NameStart == _worksheetPrefixLength
                && (_worksheetPrefixLength == 0
                    || _buffer!.AsSpan(tag.NameStart, _worksheetPrefixLength).SequenceEqual(
                        _buffer.AsSpan(_worksheetPrefixStart, _worksheetPrefixLength)));
        }
    }
}
