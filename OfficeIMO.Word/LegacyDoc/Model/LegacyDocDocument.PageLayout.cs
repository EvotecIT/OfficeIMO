namespace OfficeIMO.Word.LegacyDoc.Model {
    public sealed partial class LegacyDocDocument {
        private void ReadDopMarginSettings(byte[] tableStream, LegacyDocFib fib) {
            if (fib.FcDop < 0 || fib.FcDop > tableStream.Length - fib.LcbDop) return;

            const uint mirrorMarginsFlag = 0x00200000;
            if (fib.LcbDop >= 8) {
                uint flags = unchecked((uint)LegacyDocFib.ReadInt32(tableStream, fib.FcDop + 4));
                MirrorMargins = (flags & mirrorMarginsFlag) != 0;
            }
            // MS-DOC DopBase stores Copts60 at byte 8; bit 5 is fNoColumnBalance.
            const int compatibilityFlagsOffset = 8;
            const ushort noColumnBalanceFlag = 0x0020;
            if (fib.LcbDop >= compatibilityFlagsOffset + sizeof(ushort)) {
                ushort flags = LegacyDocFib.ReadUInt16(tableStream, fib.FcDop + compatibilityFlagsOffset);
                NoColumnBalance = (flags & noColumnBalanceFlag) != 0;
            }
            const int viewFlagsOffset = 82;
            const ushort gutterAtTopFlag = 0x8000;
            if (fib.LcbDop >= viewFlagsOffset + sizeof(ushort)) {
                ushort flags = LegacyDocFib.ReadUInt16(tableStream, fib.FcDop + viewFlagsOffset);
                GutterAtTop = (flags & gutterAtTopFlag) != 0;
            }
        }
    }
}
