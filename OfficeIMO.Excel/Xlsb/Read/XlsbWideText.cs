using System.Text;
#if NET8_0_OR_GREATER
using System.Runtime.InteropServices;
#endif

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>Decodes validated BIFF12 UTF-16LE text with Encoding.Unicode semantics.</summary>
    internal static class XlsbWideText {
        internal static string Decode(byte[] bytes, int offset, int byteCount) {
#if NET8_0_OR_GREATER
            // Text without surrogate code units cannot meet the replacement fallback, so
            // it is copied directly. Any surrogate keeps the encoding's pairing and U+FFFD rules.
            if (BitConverter.IsLittleEndian && (byteCount & 1) == 0) {
                ReadOnlySpan<char> characters = MemoryMarshal.Cast<byte, char>(bytes.AsSpan(offset, byteCount));
                if (characters.IndexOfAnyInRange('\uD800', '\uDFFF') < 0) {
                    return characters.Length == 0 ? string.Empty : new string(characters);
                }
            }
#endif
            return Encoding.Unicode.GetString(bytes, offset, byteCount);
        }

        /// <summary>
        /// Whether validated UTF-16LE text holds a surrogate code unit. Hosts that cannot
        /// test this directly report true, which keeps the full decoder.
        /// </summary>
        internal static bool ContainsSurrogate(byte[] bytes, int offset, int byteCount) {
#if NET8_0_OR_GREATER
            return !BitConverter.IsLittleEndian
                || MemoryMarshal.Cast<byte, char>(bytes.AsSpan(offset, byteCount)).IndexOfAnyInRange('\uD800', '\uDFFF') >= 0;
#else
            return true;
#endif
        }

        /// <summary>Copies text that validation proved free of surrogates, matching <see cref="Decode"/>.</summary>
        internal static string DecodeWithoutSurrogates(byte[] bytes, int offset, int byteCount) {
#if NET8_0_OR_GREATER
            return byteCount == 0 ? string.Empty : new string(MemoryMarshal.Cast<byte, char>(bytes.AsSpan(offset, byteCount)));
#else
            return Decode(bytes, offset, byteCount);
#endif
        }
    }
}
