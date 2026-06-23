using OfficeIMO.Excel.LegacyXls.Diagnostics;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffRecordReader {
        internal static IReadOnlyList<BiffRecord> ReadRecords(byte[] bytes, List<LegacyXlsImportDiagnostic> diagnostics) {
            var records = new List<BiffRecord>();
            int offset = 0;
            while (offset < bytes.Length) {
                if (offset + 4 > bytes.Length) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Error,
                        "XLS-BIFF-TRUNCATED-HEADER",
                        "The BIFF stream ended inside a record header.",
                        recordOffset: offset));
                    break;
                }

                ushort type = ReadUInt16(bytes, offset);
                ushort length = ReadUInt16(bytes, offset + 2);
                int payloadOffset = offset + 4;
                if (payloadOffset + length > bytes.Length) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Error,
                        "XLS-BIFF-TRUNCATED-PAYLOAD",
                        $"BIFF record 0x{type:X4} declares {length} payload bytes, but the stream ends early.",
                        recordOffset: offset,
                        recordType: type));
                    break;
                }

                byte[] payload = new byte[length];
                Buffer.BlockCopy(bytes, payloadOffset, payload, 0, length);
                records.Add(new BiffRecord(type, offset, payload));
                offset = payloadOffset + length;
            }

            return records;
        }

        internal static ushort ReadUInt16(byte[] bytes, int offset) {
            if (offset < 0 || offset + 2 > bytes.Length) throw new InvalidDataException("Unexpected end of BIFF record.");
            return (ushort)(bytes[offset] | (bytes[offset + 1] << 8));
        }

        internal static short ReadInt16(byte[] bytes, int offset) {
            return unchecked((short)ReadUInt16(bytes, offset));
        }

        internal static uint ReadUInt32(byte[] bytes, int offset) {
            if (offset < 0 || offset + 4 > bytes.Length) throw new InvalidDataException("Unexpected end of BIFF record.");
            return (uint)(bytes[offset]
                | (bytes[offset + 1] << 8)
                | (bytes[offset + 2] << 16)
                | (bytes[offset + 3] << 24));
        }

        internal static int ReadInt32(byte[] bytes, int offset) {
            return unchecked((int)ReadUInt32(bytes, offset));
        }

        internal static double ReadDouble(byte[] bytes, int offset) {
            if (offset < 0 || offset + 8 > bytes.Length) throw new InvalidDataException("Unexpected end of BIFF record.");
            byte[] valueBytes = new byte[8];
            Buffer.BlockCopy(bytes, offset, valueBytes, 0, valueBytes.Length);
            return BitConverter.ToDouble(valueBytes, 0);
        }
    }
}
