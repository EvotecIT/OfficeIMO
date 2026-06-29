namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocFib {
        private const int MinimumFibBytes = 0x1AA;
        private const ushort WordDocumentMagic = 0xA5EC;
        private const int FlagsOffset = 0x0A;
        private const int CcpTextOffset = 0x4C;
        private const int FcPlcfBteChpxOffset = 0xFA;
        private const int LcbPlcfBteChpxOffset = 0xFE;
        private const int FcSttbfFfnOffset = 0x112;
        private const int LcbSttbfFfnOffset = 0x116;
        private const int FcClxOffset = 0x1A2;
        private const int LcbClxOffset = 0x1A6;
        private const ushort EncryptedFlag = 0x0100;
        private const ushort OneTableStreamFlag = 0x0200;

        private LegacyDocFib(
            ushort nFib,
            bool isEncrypted,
            bool usesOneTableStream,
            int ccpText,
            int fcPlcfBteChpx,
            int lcbPlcfBteChpx,
            int fcSttbfFfn,
            int lcbSttbfFfn,
            int fcClx,
            int lcbClx) {
            NFib = nFib;
            IsEncrypted = isEncrypted;
            UsesOneTableStream = usesOneTableStream;
            CcpText = ccpText;
            FcPlcfBteChpx = fcPlcfBteChpx;
            LcbPlcfBteChpx = lcbPlcfBteChpx;
            FcSttbfFfn = fcSttbfFfn;
            LcbSttbfFfn = lcbSttbfFfn;
            FcClx = fcClx;
            LcbClx = lcbClx;
        }

        internal ushort NFib { get; }

        internal bool IsEncrypted { get; }

        internal bool UsesOneTableStream { get; }

        internal int CcpText { get; }

        internal int FcPlcfBteChpx { get; }

        internal int LcbPlcfBteChpx { get; }

        internal int FcSttbfFfn { get; }

        internal int LcbSttbfFfn { get; }

        internal int FcClx { get; }

        internal int LcbClx { get; }

        internal static bool TryRead(byte[] wordDocumentStream, out LegacyDocFib fib, out string? error) {
            fib = default;
            error = null;

            if (wordDocumentStream.Length < MinimumFibBytes) {
                error = "The WordDocument stream is too small to contain a Word 97-2003 FIB.";
                return false;
            }

            ushort magic = ReadUInt16(wordDocumentStream, 0);
            if (magic != WordDocumentMagic) {
                error = $"Unsupported WordDocument stream signature 0x{magic:X4}.";
                return false;
            }

            ushort nFib = ReadUInt16(wordDocumentStream, 0x02);
            ushort flags = ReadUInt16(wordDocumentStream, FlagsOffset);
            int ccpText = ReadInt32(wordDocumentStream, CcpTextOffset);
            int fcPlcfBteChpx = ReadInt32(wordDocumentStream, FcPlcfBteChpxOffset);
            int lcbPlcfBteChpx = ReadInt32(wordDocumentStream, LcbPlcfBteChpxOffset);
            int fcSttbfFfn = ReadInt32(wordDocumentStream, FcSttbfFfnOffset);
            int lcbSttbfFfn = ReadInt32(wordDocumentStream, LcbSttbfFfnOffset);
            int fcClx = ReadInt32(wordDocumentStream, FcClxOffset);
            int lcbClx = ReadInt32(wordDocumentStream, LcbClxOffset);

            if (ccpText < 0 || fcPlcfBteChpx < 0 || lcbPlcfBteChpx < 0 || fcSttbfFfn < 0 || lcbSttbfFfn < 0 || fcClx < 0 || lcbClx < 0) {
                error = "The FIB contains negative text or piece-table offsets.";
                return false;
            }

            fib = new LegacyDocFib(
                nFib,
                (flags & EncryptedFlag) != 0,
                (flags & OneTableStreamFlag) != 0,
                ccpText,
                fcPlcfBteChpx,
                lcbPlcfBteChpx,
                fcSttbfFfn,
                lcbSttbfFfn,
                fcClx,
                lcbClx);
            return true;
        }

        internal static ushort ReadUInt16(byte[] bytes, int offset) {
            return (ushort)(bytes[offset] | (bytes[offset + 1] << 8));
        }

        internal static int ReadInt32(byte[] bytes, int offset) {
            return bytes[offset]
                | (bytes[offset + 1] << 8)
                | (bytes[offset + 2] << 16)
                | (bytes[offset + 3] << 24);
        }
    }
}
