using System;
using System.Text;

namespace OfficeIMO.Core.Internal {
    /// <summary>Reads inert, bounded clipboard-format identification without activating an OLE object.</summary>
    internal static class OfficeOleCompoundObjectReader {
        internal static string? ClipboardFormat(byte[] bytes) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (bytes.Length < 32) return null;
            uint userLength = BitConverter.ToUInt32(bytes, 28);
            if (userLength > 4096 || userLength > bytes.Length - 32) return null;
            int offset = checked(32 + (int)userLength);
            if (offset > bytes.Length - 4) return null;
            uint length = BitConverter.ToUInt32(bytes, offset);
            // Numeric clipboard formats and absent strings are not interpreted as text.
            if (length < 1 || length > 400 || length > bytes.Length - offset - 4) return null;
            if (bytes[offset + 4 + (int)length - 1] != 0) return null;
            for (int i = 0; i < length - 1; i++) if (bytes[offset + 4 + i] == 0 || bytes[offset + 4 + i] > 127) return null;
            return Encoding.ASCII.GetString(bytes, offset + 4, (int)length - 1);
        }
    }
}
