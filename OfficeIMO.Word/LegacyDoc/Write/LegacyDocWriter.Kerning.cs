namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const ushort SprmCHpsKern = 0x484B;

        private static int ReadSupportedKerning(LegacyDocElement kern) {
            uint halfPoints = LegacyDocValues.UInt32(kern.Val) ?? 0;
            if (halfPoints > 3276) {
                throw new NotSupportedException("Native DOC kerning thresholds must be between zero and 1638 points.");
            }

            return (int)halfPoints;
        }
    }
}
