using OfficeIMO.Drawing;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    /// <summary>Bounded expanded Access designer layout. Unsupported versions remain inert opaque payloads.</summary>
    internal static partial class AccessNativeDesigner {
        internal static AccessDesignerNode? Read(byte[] bytes, int maximumNodes, CancellationToken cancellation) {
            if (bytes.Length < 10 || U16(bytes, 0) != 21 || U16(bytes, 2) != 20 || U32(bytes, 4) != 0) return null;
            int offset = 8, count = 0;
            try { AccessDesignerNode node = ReadNode(new OfficeByteView(bytes), ref offset, ref count, maximumNodes, 0, cancellation); return offset == bytes.Length ? node : null; }
            catch (InvalidDataException) { return null; }
        }
        private static AccessDesignerNode ReadNode(OfficeByteView bytes, ref int offset, ref int count, int maximumNodes, int depth, CancellationToken cancellation) {
            if (++count > maximumNodes || depth > 64 || offset > bytes.Length - 2) throw new InvalidDataException("Designer node limit or boundary exceeded.");
            ushort kind = U16(bytes, offset); offset += 2; List<AccessDesignerProperty> properties = new List<AccessDesignerProperty>(); List<AccessDesignerNode> children = new List<AccessDesignerNode>();
            while (offset < bytes.Length) {
                cancellation.ThrowIfCancellationRequested(); if (offset > bytes.Length - 4) throw new InvalidDataException("Truncated designer record.");
                uint id = U32(bytes, offset); offset += 4;
                if (id == 253) return new AccessDesignerNode(kind, properties.ToArray(), children.ToArray());
                if (id == 254) { children.Add(ReadNode(bytes, ref offset, ref count, maximumNodes, depth + 1, cancellation)); return new AccessDesignerNode(kind, properties.ToArray(), children.ToArray()); }
                if (id == 255) {
                    if (offset > bytes.Length - 2) throw new InvalidDataException("Truncated designer child group.");
                    int childCount = U16(bytes, offset); offset += 2;
                    if (childCount > maximumNodes - count) throw new InvalidDataException("Designer child group exceeds its limit.");
                    for (int child = 0; child < childCount; child++) children.Add(ReadNode(bytes, ref offset, ref count, maximumNodes, depth + 1, cancellation));
                    return new AccessDesignerNode(kind, properties.ToArray(), children.ToArray());
                }
                if (++count > maximumNodes || offset > bytes.Length - 14) throw new InvalidDataException("Designer property limit or boundary exceeded.");
                ushort code = U16(bytes, offset); uint type = U32(bytes, offset + 2), width = U32(bytes, offset + 6), size = U32(bytes, offset + 10); offset += 14;
                if (size > int.MaxValue || size > bytes.Length - offset) throw new InvalidDataException("Designer property crosses the payload boundary.");
                byte[] payload = Slice(bytes, offset, (int)size).ToArray(); offset += (int)size;
                object value = Decode(type, payload);
                properties.Add(new AccessDesignerProperty(id, code, type, width, payload, value));
            }
            if (depth != 0) throw new InvalidDataException("Designer child terminator is missing.");
            return new AccessDesignerNode(kind, properties.ToArray(), children.ToArray());
        }
        private static object Decode(uint type, byte[] bytes) {
            try {
                return type switch {
                    1 when bytes.Length == 1 => bytes[0] != 0,
                    2 when bytes.Length == 1 => bytes[0],
                    3 when bytes.Length == 2 => (short)I16(bytes, 0),
                    4 when bytes.Length == 4 => I32(bytes, 0),
                    6 when bytes.Length == 4 => BitConverter.ToSingle(bytes, 0),
                    7 when bytes.Length == 8 => F64(bytes, 0),
                    10 or 12 when (bytes.Length & 1) == 0 && bytes.Length > 0 => new System.Text.UnicodeEncoding(false, false, true).GetString(bytes).TrimEnd('\0'),
                    _ => new AccessOpaqueValue(type, bytes, "Unknown or implicit-default native designer property remains uninterpreted.")
                };
            }
            catch (System.Text.DecoderFallbackException) {
                return new AccessOpaqueValue(type, bytes, "The native designer text is not valid UTF-16; its property remains uninterpreted.");
            }
        }
    }
}
