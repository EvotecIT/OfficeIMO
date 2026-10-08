using OfficeIMO.Drawing;
using System.Text;

namespace OfficeIMO.Access {
    /// <summary>Checked Jet/ACE scalar decoding over Core's shared allocation-free byte view.</summary>
    internal static class AccessNativeBinary {
        internal static readonly Encoding Unicode = new UnicodeEncoding(false, false, true);
        internal static OfficeByteView Slice(OfficeByteView bytes, int offset, int length) {
            if (offset < 0 || length < 0 || offset > bytes.Length - length) throw new InvalidDataException("Native Access offset or length exceeds its bounded record.");
            return bytes.Slice(offset, length);
        }
        internal static ushort U16(OfficeByteView bytes, int offset) { OfficeByteView v = Slice(bytes, offset, 2); return (ushort)(v[0] | v[1] << 8); }
        internal static int I16(OfficeByteView bytes, int offset) => unchecked((short)U16(bytes, offset));
        internal static uint U32(OfficeByteView bytes, int offset) { OfficeByteView v = Slice(bytes, offset, 4); return (uint)(v[0] | v[1] << 8 | v[2] << 16 | v[3] << 24); }
        internal static int I32(OfficeByteView bytes, int offset) => unchecked((int)U32(bytes, offset));
        internal static long I64(OfficeByteView bytes, int offset) => unchecked((long)((ulong)U32(bytes, offset) | (ulong)U32(bytes, offset + 4) << 32));
        internal static double F64(OfficeByteView bytes, int offset) => BitConverter.Int64BitsToDouble(I64(bytes, offset));
        internal static string Text(OfficeByteView bytes) {
            if (bytes.Length == 0) return string.Empty;
            try {
                if (bytes.Length < 2 || bytes[0] != 0xff || bytes[1] != 0xfe) return Unicode.GetString(bytes.ToArray());
                StringBuilder text = new StringBuilder(bytes.Length); bool compressed = true; int start = 2;
                for (int end = 2; end <= bytes.Length; end++) {
                    if (end < bytes.Length && bytes[end] != 0) continue;
                    if (compressed) { for (int i = start; i < end; i++) text.Append((char)bytes[i]); }
                    else { text.Append(Unicode.GetString(Slice(bytes, start, end - start).ToArray())); }
                    compressed = !compressed; start = end + 1;
                }
                return text.ToString();
            } catch (DecoderFallbackException exception) { throw new InvalidDataException("Native Access text has invalid Unicode encoding.", exception); }
        }
        internal static string Name(OfficeByteView bytes, ref int position) {
            int length = U16(bytes, position); position = checked(position + 2);
            if (length < 2 || length > 128 || (length & 1) != 0) throw new InvalidDataException("Native Access name length is invalid.");
            string name;
            try { name = Unicode.GetString(Slice(bytes, position, length).ToArray()); }
            catch (DecoderFallbackException exception) { throw new InvalidDataException("Native Access name has invalid Unicode encoding.", exception); }
            position = checked(position + length);
            try { return AccessNamedObject.ValidateName(name); }
            catch (ArgumentException exception) { throw new InvalidDataException("Native Access name is invalid.", exception); }
        }
    }
}
