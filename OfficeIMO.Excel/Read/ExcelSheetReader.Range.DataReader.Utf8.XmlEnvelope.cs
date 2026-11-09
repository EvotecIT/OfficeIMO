#nullable enable

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
#if NET8_0_OR_GREATER
            private bool _indexedBodyHasPrefixedAttributes;
#endif
            /// <summary>
            /// Preserves XmlReader proof for enclosing markup while reusing complete
            /// indexed and character proof for a consistently qualified sheetData subtree.
            /// </summary>
            private MemoryStream CreateWorksheetXmlValidationStream(CancellationToken ct) {
#if NET8_0_OR_GREATER
                const int maximumEnvelopeBytes = 64 * 1024;
                int prefixLength = _sheetDataContentStart;
                int sheetDataTagLength = 11 + _worksheetPrefixLength;
                if (_sheetDataSupportsFastValidation
                    && prefixLength >= sheetDataTagLength
                    && _sheetDataEndTagStart >= prefixLength
                    && _sheetDataEndTagStart <= _length
                    && prefixLength <= maximumEnvelopeBytes
                    && _length - _sheetDataEndTagStart <= maximumEnvelopeBytes - prefixLength
                    && (_worksheetPrefixLength == 0
                        ? MatchesUtf8(prefixLength - sheetDataTagLength, "<sheetData>"u8)
                        : MatchesPrefixedCompactTag(prefixLength - sheetDataTagLength, "<sheetData>"u8))
                    && !_indexedBodyHasPrefixedAttributes
                    && IsCanonicalTextValid(_length, ct)
                    // Every body element uses this prefix and has no local binding.
                    // XmlReader validates its inherited binding in the intact envelope.
                    && (_worksheetPrefixLength != 0 || HasOnlyRootDefaultNamespace())) {
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
