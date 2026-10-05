using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const ushort SprmCHpsKern = 0x484B;

        private static int ReadSupportedKerning(Kern kern) {
            uint halfPoints = kern.Val?.Value ?? 0;
            if (halfPoints > 3276) {
                throw new NotSupportedException("Native DOC kerning thresholds must be between zero and 1638 points.");
            }

            return (int)halfPoints;
        }
    }
}
