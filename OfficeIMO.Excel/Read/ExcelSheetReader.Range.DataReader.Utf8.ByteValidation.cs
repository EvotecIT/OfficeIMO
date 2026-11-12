#if NET8_0_OR_GREATER
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private static bool ContainsInvalidCanonicalXmlBytes(ReadOnlySpan<byte> document) {
                int position = 0;
                if (Vector256.IsHardwareAccelerated) {
                    // Text with many entities stops here often; wider blocks halve the block count.
                    ref byte start = ref MemoryMarshal.GetReference(document);
                    var space = Vector256.Create((byte)0x20);
                    var tab = Vector256.Create((byte)'\t');
                    var cr = Vector256.Create((byte)'\r');
                    var lf = Vector256.Create((byte)'\n');
                    var amp = Vector256.Create((byte)'&');
                    var bracket = Vector256.Create((byte)']');
                    var noncharacter = Vector256.Create((byte)0xef);
                    int skipThrough = -1;
                    for (; position <= document.Length - Vector256<byte>.Count; position += Vector256<byte>.Count) {
                        var values = Vector256.LoadUnsafe(ref start, (nuint)position);
                        var controls = Vector256.LessThan(values, space)
                            & ~(Vector256.Equals(values, tab) | Vector256.Equals(values, cr) | Vector256.Equals(values, lf));
                        uint mask = (controls | Vector256.Equals(values, amp)
                            | Vector256.Equals(values, bracket) | Vector256.Equals(values, noncharacter))
                            .ExtractMostSignificantBits();
                        while (mask != 0) {
                            int index = position + BitOperations.TrailingZeroCount(mask);
                            mask &= mask - 1;
                            if (index <= skipThrough) continue;
                            if (IsInvalidCanonicalXmlByte(document, index, out int consumed)) return true;
                            skipThrough = index + consumed - 1;
                        }
                    }
                    // An entity can extend beyond the final complete block.
                    position = Math.Max(position, skipThrough + 1);
                }
                if (Vector128.IsHardwareAccelerated) {
                    ref byte start = ref MemoryMarshal.GetReference(document);
                    var space = Vector128.Create((byte)0x20);
                    var tab = Vector128.Create((byte)'\t');
                    var cr = Vector128.Create((byte)'\r');
                    var lf = Vector128.Create((byte)'\n');
                    var amp = Vector128.Create((byte)'&');
                    var bracket = Vector128.Create((byte)']');
                    var noncharacter = Vector128.Create((byte)0xef);
                    int skipThrough = -1;
                    for (; position <= document.Length - Vector128<byte>.Count; position += Vector128<byte>.Count) {
                        var values = Vector128.LoadUnsafe(ref start, (nuint)position);
                        var controls = Vector128.LessThan(values, space)
                            & ~(Vector128.Equals(values, tab) | Vector128.Equals(values, cr) | Vector128.Equals(values, lf));
                        uint mask = (controls | Vector128.Equals(values, amp)
                            | Vector128.Equals(values, bracket) | Vector128.Equals(values, noncharacter))
                            .ExtractMostSignificantBits();
                        while (mask != 0) {
                            int index = position + BitOperations.TrailingZeroCount(mask);
                            mask &= mask - 1;
                            if (index <= skipThrough) continue;
                            if (IsInvalidCanonicalXmlByte(document, index, out int consumed)) return true;
                            skipThrough = index + consumed - 1;
                        }
                    }
                    // An entity can extend beyond the final complete vector.
                    position = Math.Max(position, skipThrough + 1);
                }
                while (position < document.Length) {
                    int relative = document.Slice(position).IndexOfAny(CanonicalXmlSpecialBytes);
                    if (relative < 0) return false;
                    position += relative;
                    if (IsInvalidCanonicalXmlByte(document, position, out int consumed)) return true;
                    position += consumed;
                }
                return false;
            }

            // Bounds use the complete qualified prefix, including sequences that cross a SIMD block.
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static bool IsInvalidCanonicalXmlByte(ReadOnlySpan<byte> document, int index, out int consumed) {
                consumed = 1;
                if (document[index] == (byte)'&')
                    return !TryReadPredefinedXmlEntity(document.Slice(index), out _, out consumed);
                if (document[index] == (byte)']')
                    return index + 2 < document.Length && document[index + 1] == (byte)']'
                        && document[index + 2] == (byte)'>';
                if (document[index] == 0xef)
                    return index + 2 < document.Length && document[index + 1] == 0xbf
                        && document[index + 2] is 0xbe or 0xbf;
                return true;
            }
        }
    }
}
#endif
