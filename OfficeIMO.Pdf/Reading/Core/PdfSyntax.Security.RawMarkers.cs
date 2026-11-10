namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    [Flags]
    private enum RawSecurityName {
        None = 0,
        ByteRange = 1,
        SigFlags = 2,
        Sig = 4,
        XRef = 8,
        W = 16,
        ObjStm = 32,
        ID = 64
    }

    private readonly struct RawSecurityMarkers {
        internal RawSecurityMarkers(RawSecurityName names, int byteRangeCount) {
            Names = names;
            ByteRangeCount = byteRangeCount;
        }

        private RawSecurityName Names { get; }
        internal int ByteRangeCount { get; }
        internal bool HasSignatures => (Names & (RawSecurityName.ByteRange | RawSecurityName.SigFlags | RawSecurityName.Sig)) != 0;
        internal bool HasByteRange => ByteRangeCount != 0;
        internal bool HasXrefStreams => (Names & (RawSecurityName.XRef | RawSecurityName.W)) == (RawSecurityName.XRef | RawSecurityName.W);
        internal bool HasObjectStreams => (Names & RawSecurityName.ObjStm) != 0;
        internal bool HasTrailerId => (Names & RawSecurityName.ID) != 0;
    }

    /// <summary>Collects the conservative raw-name evidence used before the object graph is parsed.</summary>
    private static RawSecurityMarkers ScanRawSecurityMarkers(byte[] pdf, CancellationToken cancellationToken) {
        RawSecurityName names = RawSecurityName.None;
        int byteRangeCount = 0;
        const int window = 65536;
        int index = 0;
        int checkedWindowEnd = 0;
        while (index < pdf.Length) {
            int windowEnd = (int)Math.Min(pdf.Length, ((long)index / window + 1) * window);
            if (windowEnd != checkedWindowEnd) {
                cancellationToken.ThrowIfCancellationRequested();
                checkedWindowEnd = windowEnd;
            }
#if NET8_0_OR_GREATER
            int relative = pdf.AsSpan(index, windowEnd - index).IndexOf((byte)'/');
            if (relative < 0) {
                index = windowEnd;
                continue;
            }
            index += relative;
#else
            int found = Array.IndexOf(pdf, (byte)'/', index, windowEnd - index);
            if (found < 0) {
                index = windowEnd;
                continue;
            }
            index = found;
#endif

            if (index + 1 < pdf.Length) {
                switch ((char)pdf[index + 1]) {
                    case 'B' when MatchesRawPdfName(pdf, index, "ByteRange"):
                        names |= RawSecurityName.ByteRange;
                        byteRangeCount++;
                        break;
                    case 'S' when MatchesRawPdfName(pdf, index, "SigFlags"):
                        names |= RawSecurityName.SigFlags;
                        break;
                    case 'S' when MatchesRawPdfName(pdf, index, "Sig"):
                        names |= RawSecurityName.Sig;
                        break;
                    case 'X' when MatchesRawPdfName(pdf, index, "XRef"):
                        names |= RawSecurityName.XRef;
                        break;
                    case 'W' when MatchesRawPdfName(pdf, index, "W"):
                        names |= RawSecurityName.W;
                        break;
                    case 'O' when MatchesRawPdfName(pdf, index, "ObjStm"):
                        names |= RawSecurityName.ObjStm;
                        break;
                    case 'I' when MatchesRawPdfName(pdf, index, "ID"):
                        names |= RawSecurityName.ID;
                        break;
                }
            }

            index++;
        }

        cancellationToken.ThrowIfCancellationRequested();
        return new RawSecurityMarkers(names, byteRangeCount);
    }

    private static bool MatchesRawPdfName(byte[] pdf, int slashIndex, string name) {
        int afterName = slashIndex + name.Length + 1;
        if (afterName > pdf.Length) return false;
        for (int i = 0; i < name.Length; i++) {
            if (pdf[slashIndex + 1 + i] != name[i]) return false;
        }
        return afterName == pdf.Length || IsPdfDelimiter((char)pdf[afterName]) || char.IsWhiteSpace((char)pdf[afterName]);
    }
}
