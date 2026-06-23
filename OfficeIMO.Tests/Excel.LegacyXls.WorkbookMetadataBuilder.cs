using System.Text;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static partial class LegacyXlsTestWorkbookBuilder {
            internal static byte[] CreateWorkbookMetadataWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Metadata"));
                WriteRecord(stream, 0x0042, BuildUInt16Payload(1200));
                WriteRecord(stream, 0x01ba, BuildUnicodeStringPayload("ThisWorkbook"));
                WriteRecord(stream, 0x00e1, BuildUInt16Payload(1200));
                WriteRecord(stream, 0x00e2, Array.Empty<byte>());
                WriteRecord(stream, 0x005c, BuildWriteAccessPayload("OfficeIMO"));
                WriteRecord(stream, 0x0019, BuildUInt16Payload(1));
                WriteRecord(stream, 0x003d, BuildWindow1Payload());
                WriteRecord(stream, 0x0040, BuildUInt16Payload(1));
                WriteRecord(stream, 0x0863, BuildBookExtPayload());
                WriteRecord(stream, 0x008d, BuildUInt16Payload(2));
                WriteRecord(stream, 0x00da, BuildUInt16Payload(0x015d));
                WriteRecord(stream, 0x004d, BuildPrinterSettingsPayload());
                WriteRecord(stream, 0x0033, BuildUInt16Payload(2));
                WriteRecord(stream, 0x01af, BuildUInt16Payload(1));
                WriteRecord(stream, 0x01bc, BuildUInt16Payload(0x1234));
                WriteRecord(stream, 0x01bd, Array.Empty<byte>());
                WriteRecord(stream, 0x013d, BuildSheetTabIdsPayload(1, 2));
                WriteRecord(stream, 0x0160, BuildUInt16Payload(1));
                WriteRecord(stream, 0x008c, BuildCountryPayload(defaultCountryCode: 48, systemCountryCode: 1));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Workbook metadata"));
                WriteRecord(stream, 0x01ba, BuildUnicodeStringPayload("MetadataSheet"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            private static byte[] BuildWindow1Payload() {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 10);
                WriteUInt16(stream, 20);
                WriteUInt16(stream, 5000);
                WriteUInt16(stream, 4000);
                WriteUInt16(stream, 0x0038);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 600);
                return stream.ToArray();
            }

            private static byte[] BuildPrinterSettingsPayload() {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0);
                stream.WriteByte(0x01);
                stream.WriteByte(0x02);
                stream.WriteByte(0x03);
                stream.WriteByte(0x04);
                return stream.ToArray();
            }

            private static byte[] BuildBookExtPayload() {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0863);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 20);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildCountryPayload(ushort defaultCountryCode, ushort systemCountryCode) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, defaultCountryCode);
                WriteUInt16(stream, systemCountryCode);
                return stream.ToArray();
            }

            private static byte[] BuildSheetTabIdsPayload(params ushort[] tabIds) {
                using var stream = new MemoryStream();
                foreach (ushort tabId in tabIds) {
                    WriteUInt16(stream, tabId);
                }

                return stream.ToArray();
            }

            private static byte[] BuildWriteAccessPayload(string userName) {
                byte[] userNameBytes = Encoding.ASCII.GetBytes(userName);
                byte[] payload = new byte[112];
                payload[0] = checked((byte)userName.Length);
                payload[1] = 0;
                payload[2] = 0;
                Buffer.BlockCopy(userNameBytes, 0, payload, 3, userNameBytes.Length);
                for (int i = 3 + userNameBytes.Length; i < payload.Length; i++) {
                    payload[i] = 0x20;
                }

                return payload;
            }
        }
    }
}
