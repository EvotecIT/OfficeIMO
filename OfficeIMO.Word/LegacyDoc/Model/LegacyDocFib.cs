namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocFib {
        private const int MinimumFibBytes = 0x1AA;
        private const ushort WordDocumentMagic = 0xA5EC;
        private const ushort MinimumSupportedNFib = 0x00C1;
        private const int FlagsOffset = 0x0A;
        private const int CcpTextOffset = 0x4C;
        private const int CcpFtnOffset = 0x50;
        private const int CcpHddOffset = 0x54;
        private const int CcpAtnOffset = 0x5C;
        private const int CcpEdnOffset = 0x60;
        private const int CcpTxbxOffset = 0x64;
        private const int CcpHdrTxbxOffset = 0x68;
        private const int FcStshfOffset = 0xA2;
        private const int LcbStshfOffset = 0xA6;
        private const int FcPlcfBteChpxOffset = 0xFA;
        private const int LcbPlcfBteChpxOffset = 0xFE;
        private const int FcPlcfBtePapxOffset = 0x102;
        private const int LcbPlcfBtePapxOffset = 0x106;
        private const int FcSttbfFfnOffset = 0x112;
        private const int LcbSttbfFfnOffset = 0x116;
        private const int FcClxOffset = 0x1A2;
        private const int LcbClxOffset = 0x1A6;
        private const ushort FastSavedFlag = 0x0004;
        private const ushort HasPicturesFlag = 0x0008;
        private const ushort QuickSaveCountMask = 0x00F0;
        private const ushort EncryptedFlag = 0x0100;
        private const ushort OneTableStreamFlag = 0x0200;

        private LegacyDocFib(
            ushort nFib,
            bool isEncrypted,
            bool isFastSaved,
            bool hasPictures,
            int quickSaveCount,
            bool usesOneTableStream,
            int ccpText,
            int ccpFtn,
            int ccpHdd,
            int ccpAtn,
            int ccpEdn,
            int ccpTxbx,
            int ccpHdrTxbx,
            int fcStshf,
            int lcbStshf,
            int fcPlcfBteChpx,
            int lcbPlcfBteChpx,
            int fcPlcfBtePapx,
            int lcbPlcfBtePapx,
            int fcSttbfFfn,
            int lcbSttbfFfn,
            int fcClx,
            int lcbClx) {
            NFib = nFib;
            IsEncrypted = isEncrypted;
            IsFastSaved = isFastSaved;
            HasPictures = hasPictures;
            QuickSaveCount = quickSaveCount;
            UsesOneTableStream = usesOneTableStream;
            CcpText = ccpText;
            CcpFtn = ccpFtn;
            CcpHdd = ccpHdd;
            CcpAtn = ccpAtn;
            CcpEdn = ccpEdn;
            CcpTxbx = ccpTxbx;
            CcpHdrTxbx = ccpHdrTxbx;
            FcStshf = fcStshf;
            LcbStshf = lcbStshf;
            FcPlcfBteChpx = fcPlcfBteChpx;
            LcbPlcfBteChpx = lcbPlcfBteChpx;
            FcPlcfBtePapx = fcPlcfBtePapx;
            LcbPlcfBtePapx = lcbPlcfBtePapx;
            FcSttbfFfn = fcSttbfFfn;
            LcbSttbfFfn = lcbSttbfFfn;
            FcClx = fcClx;
            LcbClx = lcbClx;
        }

        internal ushort NFib { get; }

        internal bool IsEncrypted { get; }

        internal bool IsFastSaved { get; }

        internal bool HasPictures { get; }

        internal int QuickSaveCount { get; }

        internal bool UsesOneTableStream { get; }

        internal int CcpText { get; }

        internal int CcpFtn { get; }

        internal int CcpHdd { get; }

        internal int CcpAtn { get; }

        internal int CcpEdn { get; }

        internal int CcpTxbx { get; }

        internal int CcpHdrTxbx { get; }

        internal int FcStshf { get; }

        internal int LcbStshf { get; }

        internal int FcPlcfBteChpx { get; }

        internal int LcbPlcfBteChpx { get; }

        internal int FcPlcfBtePapx { get; }

        internal int LcbPlcfBtePapx { get; }

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
            if (nFib < MinimumSupportedNFib) {
                error = $"Unsupported Word FIB version 0x{nFib:X4}. OfficeIMO imports Word 97-2003 binary DOC streams with nFib 0x{MinimumSupportedNFib:X4} or newer.";
                return false;
            }

            ushort flags = ReadUInt16(wordDocumentStream, FlagsOffset);
            int ccpText = ReadInt32(wordDocumentStream, CcpTextOffset);
            int ccpFtn = ReadInt32(wordDocumentStream, CcpFtnOffset);
            int ccpHdd = ReadInt32(wordDocumentStream, CcpHddOffset);
            int ccpAtn = ReadInt32(wordDocumentStream, CcpAtnOffset);
            int ccpEdn = ReadInt32(wordDocumentStream, CcpEdnOffset);
            int ccpTxbx = ReadInt32(wordDocumentStream, CcpTxbxOffset);
            int ccpHdrTxbx = ReadInt32(wordDocumentStream, CcpHdrTxbxOffset);
            int fcStshf = ReadInt32(wordDocumentStream, FcStshfOffset);
            int lcbStshf = ReadInt32(wordDocumentStream, LcbStshfOffset);
            int fcPlcfBteChpx = ReadInt32(wordDocumentStream, FcPlcfBteChpxOffset);
            int lcbPlcfBteChpx = ReadInt32(wordDocumentStream, LcbPlcfBteChpxOffset);
            int fcPlcfBtePapx = ReadInt32(wordDocumentStream, FcPlcfBtePapxOffset);
            int lcbPlcfBtePapx = ReadInt32(wordDocumentStream, LcbPlcfBtePapxOffset);
            int fcSttbfFfn = ReadInt32(wordDocumentStream, FcSttbfFfnOffset);
            int lcbSttbfFfn = ReadInt32(wordDocumentStream, LcbSttbfFfnOffset);
            int fcClx = ReadInt32(wordDocumentStream, FcClxOffset);
            int lcbClx = ReadInt32(wordDocumentStream, LcbClxOffset);

            if (ccpText < 0
                || ccpFtn < 0
                || ccpHdd < 0
                || ccpAtn < 0
                || ccpEdn < 0
                || ccpTxbx < 0
                || ccpHdrTxbx < 0
                || fcStshf < 0
                || lcbStshf < 0
                || fcPlcfBteChpx < 0
                || lcbPlcfBteChpx < 0
                || fcPlcfBtePapx < 0
                || lcbPlcfBtePapx < 0
                || fcSttbfFfn < 0
                || lcbSttbfFfn < 0
                || fcClx < 0
                || lcbClx < 0) {
                error = "The FIB contains negative text or piece-table offsets.";
                return false;
            }

            fib = new LegacyDocFib(
                nFib,
                (flags & EncryptedFlag) != 0,
                (flags & FastSavedFlag) != 0,
                (flags & HasPicturesFlag) != 0,
                (flags & QuickSaveCountMask) >> 4,
                (flags & OneTableStreamFlag) != 0,
                ccpText,
                ccpFtn,
                ccpHdd,
                ccpAtn,
                ccpEdn,
                ccpTxbx,
                ccpHdrTxbx,
                fcStshf,
                lcbStshf,
                fcPlcfBteChpx,
                lcbPlcfBteChpx,
                fcPlcfBtePapx,
                lcbPlcfBtePapx,
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
