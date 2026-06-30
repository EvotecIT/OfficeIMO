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
        private const int FcPlcffndRefOffset = 0xAA;
        private const int LcbPlcffndRefOffset = 0xAE;
        private const int FcPlcffndTxtOffset = 0xB2;
        private const int LcbPlcffndTxtOffset = 0xB6;
        private const int FcPlcfSedOffset = 0xCA;
        private const int LcbPlcfSedOffset = 0xCE;
        private const int FcPlcfHddOffset = 0xF2;
        private const int LcbPlcfHddOffset = 0xF6;
        private const int FcPlcfBteChpxOffset = 0xFA;
        private const int LcbPlcfBteChpxOffset = 0xFE;
        private const int FcPlcfBtePapxOffset = 0x102;
        private const int LcbPlcfBtePapxOffset = 0x106;
        private const int FcSttbfFfnOffset = 0x112;
        private const int LcbSttbfFfnOffset = 0x116;
        private const int FcDopOffset = 0x192;
        private const int LcbDopOffset = 0x196;
        private const int FcPlcfendRefOffset = 0x20A;
        private const int LcbPlcfendRefOffset = 0x20E;
        private const int FcPlcfendTxtOffset = 0x212;
        private const int LcbPlcfendTxtOffset = 0x216;
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
            int fcPlcffndRef,
            int lcbPlcffndRef,
            int fcPlcffndTxt,
            int lcbPlcffndTxt,
            int fcPlcfSed,
            int lcbPlcfSed,
            int fcPlcfHdd,
            int lcbPlcfHdd,
            int fcPlcfBteChpx,
            int lcbPlcfBteChpx,
            int fcPlcfBtePapx,
            int lcbPlcfBtePapx,
            int fcSttbfFfn,
            int lcbSttbfFfn,
            int fcDop,
            int lcbDop,
            int fcPlcfendRef,
            int lcbPlcfendRef,
            int fcPlcfendTxt,
            int lcbPlcfendTxt,
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
            FcPlcffndRef = fcPlcffndRef;
            LcbPlcffndRef = lcbPlcffndRef;
            FcPlcffndTxt = fcPlcffndTxt;
            LcbPlcffndTxt = lcbPlcffndTxt;
            FcPlcfSed = fcPlcfSed;
            LcbPlcfSed = lcbPlcfSed;
            FcPlcfHdd = fcPlcfHdd;
            LcbPlcfHdd = lcbPlcfHdd;
            FcPlcfBteChpx = fcPlcfBteChpx;
            LcbPlcfBteChpx = lcbPlcfBteChpx;
            FcPlcfBtePapx = fcPlcfBtePapx;
            LcbPlcfBtePapx = lcbPlcfBtePapx;
            FcSttbfFfn = fcSttbfFfn;
            LcbSttbfFfn = lcbSttbfFfn;
            FcDop = fcDop;
            LcbDop = lcbDop;
            FcPlcfendRef = fcPlcfendRef;
            LcbPlcfendRef = lcbPlcfendRef;
            FcPlcfendTxt = fcPlcfendTxt;
            LcbPlcfendTxt = lcbPlcfendTxt;
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

        internal int FcPlcffndRef { get; }

        internal int LcbPlcffndRef { get; }

        internal int FcPlcffndTxt { get; }

        internal int LcbPlcffndTxt { get; }

        internal int FcPlcfSed { get; }

        internal int LcbPlcfSed { get; }

        internal int FcPlcfHdd { get; }

        internal int LcbPlcfHdd { get; }

        internal int FcPlcfBteChpx { get; }

        internal int LcbPlcfBteChpx { get; }

        internal int FcPlcfBtePapx { get; }

        internal int LcbPlcfBtePapx { get; }

        internal int FcSttbfFfn { get; }

        internal int LcbSttbfFfn { get; }

        internal int FcDop { get; }

        internal int LcbDop { get; }

        internal int FcPlcfendRef { get; }

        internal int LcbPlcfendRef { get; }

        internal int FcPlcfendTxt { get; }

        internal int LcbPlcfendTxt { get; }

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
            int fcPlcffndRef = ReadInt32(wordDocumentStream, FcPlcffndRefOffset);
            int lcbPlcffndRef = ReadInt32(wordDocumentStream, LcbPlcffndRefOffset);
            int fcPlcffndTxt = ReadInt32(wordDocumentStream, FcPlcffndTxtOffset);
            int lcbPlcffndTxt = ReadInt32(wordDocumentStream, LcbPlcffndTxtOffset);
            int fcPlcfSed = ReadInt32(wordDocumentStream, FcPlcfSedOffset);
            int lcbPlcfSed = ReadInt32(wordDocumentStream, LcbPlcfSedOffset);
            int fcPlcfHdd = ReadInt32(wordDocumentStream, FcPlcfHddOffset);
            int lcbPlcfHdd = ReadInt32(wordDocumentStream, LcbPlcfHddOffset);
            int fcPlcfBteChpx = ReadInt32(wordDocumentStream, FcPlcfBteChpxOffset);
            int lcbPlcfBteChpx = ReadInt32(wordDocumentStream, LcbPlcfBteChpxOffset);
            int fcPlcfBtePapx = ReadInt32(wordDocumentStream, FcPlcfBtePapxOffset);
            int lcbPlcfBtePapx = ReadInt32(wordDocumentStream, LcbPlcfBtePapxOffset);
            int fcSttbfFfn = ReadInt32(wordDocumentStream, FcSttbfFfnOffset);
            int lcbSttbfFfn = ReadInt32(wordDocumentStream, LcbSttbfFfnOffset);
            int fcDop = ReadInt32(wordDocumentStream, FcDopOffset);
            int lcbDop = ReadInt32(wordDocumentStream, LcbDopOffset);
            int fcPlcfendRef = ReadOptionalInt32(wordDocumentStream, FcPlcfendRefOffset);
            int lcbPlcfendRef = ReadOptionalInt32(wordDocumentStream, LcbPlcfendRefOffset);
            int fcPlcfendTxt = ReadOptionalInt32(wordDocumentStream, FcPlcfendTxtOffset);
            int lcbPlcfendTxt = ReadOptionalInt32(wordDocumentStream, LcbPlcfendTxtOffset);
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
                || fcPlcffndRef < 0
                || lcbPlcffndRef < 0
                || fcPlcffndTxt < 0
                || lcbPlcffndTxt < 0
                || fcPlcfSed < 0
                || lcbPlcfSed < 0
                || fcPlcfHdd < 0
                || lcbPlcfHdd < 0
                || fcPlcfBteChpx < 0
                || lcbPlcfBteChpx < 0
                || fcPlcfBtePapx < 0
                || lcbPlcfBtePapx < 0
                || fcSttbfFfn < 0
                || lcbSttbfFfn < 0
                || fcDop < 0
                || lcbDop < 0
                || fcPlcfendRef < 0
                || lcbPlcfendRef < 0
                || fcPlcfendTxt < 0
                || lcbPlcfendTxt < 0
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
                fcPlcffndRef,
                lcbPlcffndRef,
                fcPlcffndTxt,
                lcbPlcffndTxt,
                fcPlcfSed,
                lcbPlcfSed,
                fcPlcfHdd,
                lcbPlcfHdd,
                fcPlcfBteChpx,
                lcbPlcfBteChpx,
                fcPlcfBtePapx,
                lcbPlcfBtePapx,
                fcSttbfFfn,
                lcbSttbfFfn,
                fcDop,
                lcbDop,
                fcPlcfendRef,
                lcbPlcfendRef,
                fcPlcfendTxt,
                lcbPlcfendTxt,
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

        private static int ReadOptionalInt32(byte[] bytes, int offset) {
            return offset + 4 <= bytes.Length ? ReadInt32(bytes, offset) : 0;
        }
    }
}
