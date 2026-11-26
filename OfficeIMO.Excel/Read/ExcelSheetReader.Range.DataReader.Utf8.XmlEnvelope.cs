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
            private Stream CreateWorksheetXmlValidationStream(CancellationToken ct) {
                Stream? envelope = TryCreateWorksheetXmlEnvelope(ct);
                if (envelope != null) return envelope;
                if (_windowedPart) return OpenWindowedValidationStream(ct);
                return new MemoryStream(_buffer!, 0, _length, writable: false, publiclyVisible: false);
            }

            /// <summary>
            /// The worksheet without its sheetData body when every row passed the fast checks, or null when the whole
            /// part needs the XML reader.
            /// </summary>
            private Stream? TryCreateWorksheetXmlEnvelope(CancellationToken ct) {
#if NET8_0_OR_GREATER
                const int maximumEnvelopeBytes = 64 * 1024;
                int prefixLength = _sheetDataContentStart;
                int sheetDataTagLength = 11 + _worksheetPrefixLength;
                if (_sheetDataSupportsFastValidation
                    && prefixLength >= sheetDataTagLength
                    && _sheetDataEndTagStart >= prefixLength
                    && _sheetDataEndTagStart <= _length
                    && IsSheetDataEndTagAt(_sheetDataEndTagStart)
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
                return null;
            }

#if NET8_0_OR_GREATER
            // A window move or delivery restart can leave the recorded offset pointing at other bytes.
            private bool IsSheetDataEndTagAt(int position) =>
                position + 12 + _worksheetPrefixLength <= _length
                && _buffer![position] == (byte)'<' && _buffer[position + 1] == (byte)'/'
                && _buffer.AsSpan(position + 2 + _worksheetPrefixLength, 10).SequenceEqual("sheetData>"u8);
#endif

            /// <summary>A windowed source no longer holds the whole part, so complete XML proof reads it again.</summary>
            private Stream OpenWindowedValidationStream(CancellationToken ct) {
                if (!_owner._partBufferReader!.TryOpenExactPartStream(_owner._worksheetPartName, ct, out Stream? stream)) {
                    throw new InvalidDataException($"Worksheet '{_owner._sheetName}' could not be reopened for validation.");
                }
                return stream!;
            }
        }
    }
}
