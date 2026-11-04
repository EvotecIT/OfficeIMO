#nullable enable

using System.Threading;
#if NET8_0_OR_GREATER
using System.Buffers;
using System.Text.Unicode;
#endif

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
#if NET8_0_OR_GREATER
            private static readonly SearchValues<byte> NonCanonicalAsciiBytes = CreateNonCanonicalAsciiBytes();
            private int _validatedCanonicalTextLength;
            private bool _canonicalTextNeedsXmlReader;
            private bool _canonicalTextHasNonSimpleBytes;

            private static SearchValues<byte> CreateNonCanonicalAsciiBytes() {
                Span<byte> bytes = stackalloc byte[256];
                int count = 0;
                for (int value = 0; value <= 255; value++) {
                    if (value >= 128 || value < 32 && value is not 9 and not 10 and not 13
                        || value is '&' or ']' or 13) bytes[count++] = (byte)value;
                }
                return SearchValues.Create(bytes.Slice(0, count));
            }

            /// <summary>
            /// Qualifies initialized text once. Published row-token prefixes end in
            /// ASCII, so a valid UTF-8 code point cannot straddle their boundaries.
            /// The same boundary also cannot split either forbidden XML sequence.
            /// Non-ASCII and special bytes retain complete UTF-8/XML-character proof;
            /// entities or richer content still require the ordinary XmlReader proof.
            /// </summary>
            private bool IsCanonicalTextValid(int length, CancellationToken ct) {
                ct.ThrowIfCancellationRequested();
                if (_canonicalTextNeedsXmlReader) return false;
                ReadOnlySpan<byte> text = _buffer!.AsSpan(_validatedCanonicalTextLength,
                    length - _validatedCanonicalTextLength);
                if (text.IndexOfAny(NonCanonicalAsciiBytes) >= 0) {
                    _canonicalTextHasNonSimpleBytes = true;
                    if (!Utf8.IsValid(text) || ContainsInvalidCanonicalXmlBytes(text)) {
                        _canonicalTextNeedsXmlReader = true;
                        return false;
                    }
                }
                _validatedCanonicalTextLength = length;
                ct.ThrowIfCancellationRequested();
                return true;
            }
#endif

            /// <summary>Overlaps character proof with inflation; publication still requires complete XML proof.</summary>
            private void ValidateProducedCanonicalText(CancellationToken ct) {
#if NET8_0_OR_GREATER
                if (_sheetDataSupportsFastValidation) _ = IsCanonicalTextValid(_length, ct);
#endif
            }
        }
    }
}
