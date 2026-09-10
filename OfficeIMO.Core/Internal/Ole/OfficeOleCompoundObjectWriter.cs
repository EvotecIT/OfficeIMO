using System;
using System.IO;
using System.Text;

namespace OfficeIMO.Core.Internal {
    /// <summary>Writes the inert OLE CompObj identification stream defined by MS-OLEDS 2.3.8.</summary>
    internal static class OfficeOleCompoundObjectWriter {
        internal static byte[] Write(Guid classId, string userType, string clipboardFormat, string programId) {
            using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
            // Conventional header values; MS-OLEDS reserves these bytes for the producing application.
            writer.Write(0xfffe0001u); writer.Write(0x00000a03u); writer.Write(uint.MaxValue); writer.Write(classId.ToByteArray());
            WriteAnsi(writer, userType, 4096); WriteAnsi(writer, clipboardFormat, 400); WriteAnsi(writer, programId, 40);
            writer.Write(0x71b239f4u); writer.Write(0); writer.Write(0); writer.Write(0);
            return stream.ToArray();
        }
        private static void WriteAnsi(BinaryWriter writer, string value, int maximum) {
            if (value == null) throw new ArgumentNullException(nameof(value));
            if (value.Length >= maximum) throw new ArgumentOutOfRangeException(nameof(value));
            foreach (char character in value) if (character == 0 || character > 127)
                throw new ArgumentException("OLE identification strings must be ASCII without embedded NUL characters.", nameof(value));
            byte[] bytes = Encoding.ASCII.GetBytes(value); writer.Write(bytes.Length + 1); writer.Write(bytes); writer.Write((byte)0);
        }
    }
}
