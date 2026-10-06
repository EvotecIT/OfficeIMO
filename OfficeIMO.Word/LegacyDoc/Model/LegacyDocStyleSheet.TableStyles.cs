namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed partial class LegacyDocStyleSheet {
        internal LegacyDocTableBorders ResolveTableBorders(ushort? styleIndex) {
            LegacyDocTableBorders borders = default;
            var visited = new HashSet<ushort>();
            while (styleIndex.HasValue && visited.Add(styleIndex.Value)
                && _tableStyles.TryGetValue(styleIndex.Value, out var style)) {
                borders = borders.WithDefaults(style.Borders);
                styleIndex = style.BasedOnStyleIndex;
            }
            return borders;
        }

        private static bool TryReadTableStyle(byte[] bytes, int offset, int count, int baseSize,
            out LegacyDocTableStyle? style) {
            style = null;
            if (count < baseSize || (LegacyDocFib.ReadUInt16(bytes, offset + 2) & 0x0F) != 3) return false;
            int end = offset + count;
            ushort parent = (ushort)(LegacyDocFib.ReadUInt16(bytes, offset + 2) >> 4);
            int upxCount = LegacyDocFib.ReadUInt16(bytes, offset + 4) & 0x0F;
            string? name = ReadXstz(bytes, offset + baseSize, end, out int upxOffset);
            if (string.IsNullOrWhiteSpace(name) || upxCount == 0) return false;
            if ((upxOffset & 1) != 0) upxOffset++;
            if (upxOffset + 2 > end) return false;
            int length = LegacyDocFib.ReadUInt16(bytes, upxOffset);
            upxOffset += 2;
            if (upxOffset + length > end) return false;
            style = new LegacyDocTableStyle(parent == 0x0FFF ? null : parent,
                LegacyDocParagraphFormattingReader.ReadTableStyleBorders(bytes, upxOffset, length));
            return true;
        }

        private sealed class LegacyDocTableStyle {
            internal LegacyDocTableStyle(ushort? basedOnStyleIndex, LegacyDocTableBorders borders) {
                BasedOnStyleIndex = basedOnStyleIndex;
                Borders = borders;
            }
            internal ushort? BasedOnStyleIndex { get; }
            internal LegacyDocTableBorders Borders { get; }
        }
    }
}
