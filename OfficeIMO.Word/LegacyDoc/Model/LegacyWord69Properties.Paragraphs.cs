namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed partial class LegacyWord69Properties {
        /// <summary>Converts native tabs and borders into the common paragraph/table property grammar.</summary>
        private bool AppendParagraphStructure(List<byte> output, byte operation, LegacyWord69Bytes operand) {
            if (operation is 15 or 23) {
                int deleted = operand.Length == 0 ? -1 : operand[0];
                int addedIndex = 1 + deleted * 2;
                if (deleted < 0 || addedIndex >= operand.Length
                    || (operation == 23 && deleted != 0)
                    || addedIndex + 1 + 3 * operand[addedIndex] != operand.Length) {
                    HasUnsupportedFormatting = true;
                    return true;
                }
                int descriptors = addedIndex + 1 + 2 * operand[addedIndex];
                for (int index = descriptors; index < operand.Length; index++) {
                    int alignment = operand[index] & 7;
                    int leader = (operand[index] >> 3) & 7;
                    if (alignment is 5 or 7 || leader > 4 || (operand[index] & 0xC0) != 0) {
                        HasUnsupportedFormatting = true;
                    }
                }
                AppendCode(output, 0xC60D);
                output.Add((byte)operand.Length);
                for (int index = 0; index < operand.Length; index++) output.Add(operand[index]);
                return true;
            }
            if (operation is >= 38 and <= 42) {
                AppendCode(output, (ushort)(0x6424 + operation - 38));
                AppendBorder(output, operand);
                return true;
            }
            if (operation == 187) {
                AppendCode(output, 0xD605);
                output.Add(24);
                for (int index = 0; index < 6; index++) AppendBorder(output, operand.Slice(index * 2, 2));
                return true;
            }
            return false;
        }

        private void AppendBorder(List<byte> output, LegacyWord69Bytes operand) {
            ushort native = LegacyWord69Reader.U16(operand, 0);
            int width = native & 7;
            int type = (native >> 3) & 3;
            if (width > 5) {
                type = width;
                width = 1;
            }
            int color = (native >> 6) & 31;
            if ((native & 0x20) != 0 || color > 16 || type == 2) HasUnsupportedFormatting = true;
            output.Add((byte)(width * 6));
            output.Add((byte)type);
            output.Add((byte)color);
            output.Add((byte)(native >> 11));
        }
    }
}
