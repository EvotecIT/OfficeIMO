#nullable enable

using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
#if NET8_0_OR_GREATER
            private bool _indexedBodyHasPrefixedAttributes;
#endif
            /// <summary>
            /// Preserves XmlReader proof for enclosing markup while reusing complete
            /// indexed and character proof for an unprefixed sheetData subtree.
            /// </summary>
            private MemoryStream CreateWorksheetXmlValidationStream(CancellationToken ct) {
#if NET8_0_OR_GREATER
                const int maximumEnvelopeBytes = 64 * 1024;
                int prefixLength = _sheetDataContentStart;
                if (_sheetDataSupportsFastValidation
                    && _worksheetPrefixLength == 0
                    && prefixLength >= 11
                    && _sheetDataEndTagStart >= prefixLength
                    && _sheetDataEndTagStart <= _length
                    && prefixLength <= maximumEnvelopeBytes
                    && _length - _sheetDataEndTagStart <= maximumEnvelopeBytes - prefixLength
                    && MatchesUtf8(prefixLength - 11, "<sheetData>"u8)
                    && !_indexedBodyHasPrefixedAttributes
                    && IsCanonicalTextValid(_length, ct)
                    && HasOnlyRootDefaultNamespace()) {
                    int suffixLength = _length - _sheetDataEndTagStart;
                    byte[] envelope = new byte[prefixLength + suffixLength];
                    _buffer!.AsSpan(0, prefixLength).CopyTo(envelope);
                    _buffer.AsSpan(_sheetDataEndTagStart, suffixLength).CopyTo(envelope.AsSpan(prefixLength));
                    return new MemoryStream(envelope, writable: false);
                }
#endif
                return new MemoryStream(_buffer!, 0, _length, writable: false, publiclyVisible: false);
            }
        }
    }
}
