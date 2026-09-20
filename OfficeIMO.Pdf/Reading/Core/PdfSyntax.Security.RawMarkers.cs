using System.Threading;

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
    private static RawSecurityMarkers ScanRawSecurityMarkers(string text, CancellationToken cancellationToken) {
        RawSecurityName names = RawSecurityName.None;
        int byteRangeCount = 0;
        bool cancellable = cancellationToken.CanBeCanceled;
        int index = 0;
        while (index < text.Length) {
            if (cancellable) {
                if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (text[index] != '/') {
                    index++;
                    continue;
                }
            } else {
                index = text.IndexOf('/', index);
                if (index < 0) break;
            }

            if (index + 1 < text.Length) {
                switch (text[index + 1]) {
                    case 'B' when MatchesRawPdfName(text, index, "ByteRange"):
                        names |= RawSecurityName.ByteRange;
                        byteRangeCount++;
                        break;
                    case 'S' when MatchesRawPdfName(text, index, "SigFlags"):
                        names |= RawSecurityName.SigFlags;
                        break;
                    case 'S' when MatchesRawPdfName(text, index, "Sig"):
                        names |= RawSecurityName.Sig;
                        break;
                    case 'X' when MatchesRawPdfName(text, index, "XRef"):
                        names |= RawSecurityName.XRef;
                        break;
                    case 'W' when MatchesRawPdfName(text, index, "W"):
                        names |= RawSecurityName.W;
                        break;
                    case 'O' when MatchesRawPdfName(text, index, "ObjStm"):
                        names |= RawSecurityName.ObjStm;
                        break;
                    case 'I' when MatchesRawPdfName(text, index, "ID"):
                        names |= RawSecurityName.ID;
                        break;
                }
            }

            index++;
        }

        cancellationToken.ThrowIfCancellationRequested();
        return new RawSecurityMarkers(names, byteRangeCount);
    }

    private static bool MatchesRawPdfName(string text, int slashIndex, string name) {
        int afterName = slashIndex + name.Length + 1;
        return afterName <= text.Length &&
            string.CompareOrdinal(text, slashIndex + 1, name, 0, name.Length) == 0 &&
            (afterName == text.Length || IsPdfDelimiter(text[afterName]) || char.IsWhiteSpace(text[afterName]));
    }
}
