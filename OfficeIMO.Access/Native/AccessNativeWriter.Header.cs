using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using System.Text;
namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeWriter {
        private static byte[] Header(bool ace, DateTime createdAt) {
            byte[] page = new byte[PageSize]; page[1] = 1;
            Encoding.ASCII.GetBytes(ace ? "Standard ACE DB" : "Standard Jet DB").CopyTo(page, 4);
            page[20] = (byte)(ace ? 2 : 1);
            // Logical fixed-root declarations and the model's creation date for the unprotected profile.
            U32(page, 24, 0x100); U32(page, 28, 0x101);
            for (int root = 2; root <= 5; root++) U32(page, 24 + root * 4, (uint)root);
            U16(page, 60, 1252); U32(page, 110, 1033);
            double date = createdAt.ToOADate();
            ulong bits = unchecked((ulong)BitConverter.DoubleToInt64Bits(date)); U32(page, 114, (uint)bits); U32(page, 118, (uint)(bits >> 32));
            byte[] dateMask = new byte[4]; U32(dateMask, 0, (uint)(int)date);
            for (int i = 0; i < 40; i++) page[66 + i] = dateMask[i % 4];
            // Observed unprotected header field at 0x6A; further security-profile semantics remain unqualified.
            U32(page, 106, 4518);
            U32(page, 152, 1620); Encoding.ASCII.GetBytes("4.0").CopyTo(page, 156);
            // Empty Jet/ACE user-slot declarations; zero-filled slots are interpreted as occupied by DAO.
            page[3584] = 1; for (int slot = 3585; slot < PageSize; slot += 2) page[slot] = 1;
            // Jet's fixed header obfuscation is separate from database password/encryption protection.
            byte[] pad = Rc4Pad(new byte[] { 0xc7, 0xda, 0x39, 0x6b }, 128);
            for (int n = 0; n < pad.Length; n++) page[24 + n] ^= pad[n];
            return page;
        }

        private static byte[] SidMask(byte[] maskedHeader) {
            // Jet/ACE SID fields use their own RC4 key folded from the clear header, including its date and password region.
            // Physical contract reference: Jackcess Encrypt's public SidRemasker / JetPasswordHandler documentation.
            byte[] header = (byte[])maskedHeader.Clone(); byte[] headerMask = Rc4Pad(new byte[] { 0xc7, 0xda, 0x39, 0x6b }, 128);
            for (int i = 0; i < headerMask.Length; i++) header[24 + i] ^= headerMask[i];
            OfficeByteView view = new OfficeByteView(header);
            uint folded = AccessNativeBinary.U32(view, 114);
            int dateMask = (int)AccessNativeBinary.F64(view, 114);
            for (int i = 0; i < 40; i++) {
                int position = i * 2; byte value = header[66 + position];
                if (position < 40) value ^= (byte)(dateMask >> ((position % 4) * 8));
                folded ^= (uint)value << (i % 24);
            }
            byte[] key = new byte[4]; U32(key, 0, folded); return Rc4Pad(key, 2);
        }

        private static byte[] Rc4Pad(byte[] key, int length) {
            OfficeRc4Transform transform = new OfficeIMO.Core.Internal.OfficeRc4Transform(key);
            byte[] bytes = new byte[length];
            for (int i = 0; i < length; i++) bytes[i] = transform.NextByte();
            return bytes;
        }
    }
}
