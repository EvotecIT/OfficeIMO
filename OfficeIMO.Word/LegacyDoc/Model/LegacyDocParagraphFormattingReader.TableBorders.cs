namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static partial class LegacyDocParagraphFormattingReader {
        private const ushort SprmTIstd = 0x563A;
        private const ushort SprmTTableBorders = 0xD613;
        private const ushort SprmTTableBorders80 = 0xD605;
        private const ushort SprmTBrcTopCv = 0xD61A;
        private const ushort SprmTBrcRightCv = 0xD61D;
        private const ushort SprmTSetBrc = 0xD62F;

        private static bool TryReadTableCellBorderRange(byte[] bytes, int offset, int end,
            int cellCount, ref IReadOnlyList<LegacyDocTableCellBorders>? cellBorders) {
            if (end - offset < 14 || bytes[offset + 2] != 11) return false;
            int first = bytes[offset + 3];
            int limit = bytes[offset + 4];
            byte edges = bytes[offset + 5];
            if (first >= cellCount || first > limit || limit > cellCount || (edges & 0xC0) != 0) return false;
            if (first == limit) return true;
            // The model preserves the four cell sides. The Data-stream projection
            // guard continues to hold records containing diagonal borders.
            LegacyDocTableCellBorder border = ReadTableBrc(bytes, offset + 6);
            var result = new LegacyDocTableCellBorders[Math.Max(cellCount, cellBorders?.Count ?? 0)];
            for (int cell = 0; cell < result.Length; cell++) {
                var previous = cellBorders != null && cell < cellBorders.Count ? cellBorders[cell] : default;
                result[cell] = cell < first || cell >= limit ? previous
                    : new LegacyDocTableCellBorders((edges & 1) != 0 ? border : previous.Top,
                        (edges & 2) != 0 ? border : previous.Left,
                        (edges & 4) != 0 ? border : previous.Bottom,
                        (edges & 8) != 0 ? border : previous.Right);
            }
            cellBorders = result;
            return true;
        }

        private static bool TryReadTableCellBorderColors(byte[] bytes, int offset, int end, ushort sprm,
            ref IReadOnlyList<LegacyDocTableCellBorders>? cellBorders, out int operandLength) {
            operandLength = 0;
            if (offset + 3 > end) return false;
            int count = bytes[offset + 2];
            if (count % 4 != 0 || offset + 3 + count > end) return false;
            operandLength = count + 1;
            int cells = count / 4;
            var result = new LegacyDocTableCellBorders[Math.Max(cells, cellBorders?.Count ?? 0)];
            for (int cell = 0; cell < result.Length; cell++) {
                var original = cellBorders != null && cell < cellBorders.Count ? cellBorders[cell] : default;
                if (cell >= cells) { result[cell] = original; continue; }
                int at = offset + 3 + cell * 4;
                bool noBorder = bytes[at] == 0xFF && bytes[at + 1] == 0xFF
                    && bytes[at + 2] == 0xFF && bytes[at + 3] == 0xFF;
                string? color = bytes[at + 3] == 0xFF ? null
                    : bytes[at].ToString("X2") + bytes[at + 1].ToString("X2") + bytes[at + 2].ToString("X2");
                LegacyDocTableCellBorder Color(LegacyDocTableCellBorder border) => noBorder
                    ? new LegacyDocTableCellBorder(LegacyDocTableCellBorderStyle.ExplicitNone, null, 0, 0)
                    : new LegacyDocTableCellBorder(border.Style, color, border.SizeEighthPoints, border.SpacePoints);
                int edge = sprm - SprmTBrcTopCv;
                result[cell] = new LegacyDocTableCellBorders(edge == 0 ? Color(original.Top) : original.Top,
                    edge == 1 ? Color(original.Left) : original.Left, edge == 2 ? Color(original.Bottom) : original.Bottom,
                    edge == 3 ? Color(original.Right) : original.Right);
            }
            cellBorders = result;
            return true;
        }

        internal static LegacyDocTableBorders ReadTableStyleBorders(byte[] bytes, int offset, int count) {
            int end = offset + count;
            LegacyDocTableBorders borders = default;
            while (offset + 2 <= end) {
                ushort sprm = LegacyDocFib.ReadUInt16(bytes, offset);
                if (!TryGetSprmOperandLength(bytes, offset, end, out int length)) break;
                if (sprm == SprmTTableBorders || sprm == SprmTTableBorders80) {
                    // UpxTapx ignores sprmTIstd; only the declared parent participates in inheritance.
                    if (TryReadTableBorders(bytes, offset, end, sprm == SprmTTableBorders80, out var value)) borders = value;
                }
                offset += 2 + length;
            }
            return borders;
        }

        private static bool TryReadTableBorders(byte[] bytes, int offset, int end, bool legacy, out LegacyDocTableBorders borders) {
            borders = default;
            int size = legacy ? 4 : 8;
            if (offset + 3 + 6 * size > end || bytes[offset + 2] != 6 * size) return false;
            LegacyDocTableCellBorder Read(int edge) {
                int at = offset + 3 + edge * size;
                if (!legacy) return ReadTableBrc(bytes, at);
                // Whole-row no-border defaults suppress inheritance; zero TC80 cell borders do not.
                if (bytes[at + 1] == 0) return new LegacyDocTableCellBorder(LegacyDocTableCellBorderStyle.ExplicitNone, null, 0, 0);
                return ReadBrc80(bytes, at);
            }
            borders = new LegacyDocTableBorders(Read(0), Read(1), Read(2), Read(3), Read(4), Read(5));
            return true;
        }

        private static LegacyDocTableCellBorder ReadTableBrc(byte[] bytes, int offset) {
            // BrcMayBeNil uses a nil final DWORD. Nil and brcType=0 both suppress an inherited border.
            bool nil = bytes[offset + 4] == 0xFF && bytes[offset + 5] == 0xFF
                && bytes[offset + 6] == 0xFF && bytes[offset + 7] == 0xFF;
            byte type = bytes[offset + 5];
            if (nil || type == 0) return new LegacyDocTableCellBorder(LegacyDocTableCellBorderStyle.ExplicitNone, null, 0, 0);
            var style = MapBrc80BorderStyle(type);
            if (style == LegacyDocTableCellBorderStyle.None) return default;
            string? color = bytes[offset + 3] == 0xFF ? null
                : bytes[offset].ToString("X2") + bytes[offset + 1].ToString("X2") + bytes[offset + 2].ToString("X2");
            return new LegacyDocTableCellBorder(style, color, Math.Max(2, (int)bytes[offset + 4]), bytes[offset + 6] & 0x1F);
        }
    }
}
