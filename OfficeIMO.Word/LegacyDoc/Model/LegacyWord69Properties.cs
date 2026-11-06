namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Maps supported one-byte Word 6/95 property operands to the common legacy DOC format model.</summary>
    internal sealed class LegacyWord69Properties {
        internal bool HasUnsupportedFormatting { get; set; }
        internal int DefaultFontIndex { get; set; }
        internal Dictionary<ushort, int> StyleFonts { get; } = new();
        internal Dictionary<ushort, ushort> StyleLanguages { get; } = new();

        internal LegacyDocCharacterFormat ReadCharacter(LegacyWord69Bytes data, IReadOnlyList<string> fonts, out int? fontIndex, out ushort? languageId) {
            byte[] normalized = Normalize(data, paragraph: false, out fontIndex, out languageId);
            return LegacyDocCharacterFormattingReader.ReadGrpprl(normalized, 0, normalized.Length, fonts);
        }

        internal LegacyDocParagraphFormat ReadParagraph(LegacyWord69Bytes data, ushort styleIndex) {
            byte[] normalized = Normalize(data, paragraph: true, out _, out _, styleIndex);
            return LegacyDocParagraphFormattingReader.ReadGrpprl(normalized, 0, normalized.Length);
        }

        private byte[] Normalize(LegacyWord69Bytes data, bool paragraph, out int? fontIndex, out ushort? languageId, ushort styleIndex = 0) {
            fontIndex = null;
            languageId = null;
            var output = new List<byte>(data.Length);
            if (paragraph) {
                AppendCode(output, 0x4600);
                AppendWord(output, styleIndex);
            }

            while (!data.IsEmpty) {
                byte operation = data[0];
                data = data.Slice(1);
                if (operation == 0) continue;
                if (!TryReadOperand(operation, ref data, out var operand)) {
                    HasUnsupportedFormatting = true;
                    break;
                }

                if (paragraph && operation == 190) {
                    AppendTableDefinition(output, operand);
                    continue;
                }

                ushort code = paragraph ? ParagraphCode(operation) : CharacterCode(operation);
                if (code == 0) {
                    if (!paragraph && operation == 117) continue;
                    HasUnsupportedFormatting = true;
                    continue;
                }

                AppendCode(output, code);
                for (int i = 0; i < operand.Length; i++) output.Add(operand[i]);
                if (!paragraph && operation == 93) fontIndex = LegacyWord69Reader.U16(operand, 0);
                if (!paragraph && operation == 97) languageId = LegacyWord69Reader.U16(operand, 0);
            }

            return output.ToArray();
        }

        private static ushort ParagraphCode(byte operation) => operation switch {
            2 => 0x4600,    // paragraph style
            5 => 0x2403,    // justification
            7 => 0x2405,    // keep together
            8 => 0x2406,    // keep with next
            9 => 0x2407,    // page break before
            14 => 0x240C,   // no line numbers
            16 => 0x840E,   // right indent
            17 => 0x840F,   // left indent
            19 => 0x8411,   // first-line indent
            20 => 0x6412,   // line spacing
            21 => 0xA413,   // space before
            22 => 0xA414,   // space after
            24 => 0x2416,   // in table
            25 => 0x2417,   // row terminator
            44 => 0x242A,   // suppress hyphenation
            47 => 0x442D,   // shading
            51 => 0x2431,   // widow control
            182 => 0x548A,  // table alignment
            185 => 0x3403,  // disallow row split
            186 => 0x3404,  // repeated heading row
            189 => 0x9407,  // row height
            _ => 0
        };

        private static ushort CharacterCode(byte operation) => operation switch {
            65 => 0x0800,   // deleted revision
            66 => 0x0801,   // inserted revision
            85 => 0x0835,   // bold
            86 => 0x0836,   // italic
            87 => 0x0837,   // strike
            88 => 0x0838,   // outline
            89 => 0x0839,   // shadow
            90 => 0x083A,   // small caps
            91 => 0x083B,   // caps
            92 => 0x083C,   // hidden
            93 => 0x4A4F,   // font index
            94 => 0x2A3E,   // underline
            96 => 0x8840,   // spacing
            97 => 0x486D,   // language
            98 => 0x2A42,   // palette color
            99 => 0x4A43,   // font size in half-points
            104 => 0x2A48,  // superscript and subscript
            107 => 0x484B,  // kerning
            _ => 0
        };

        private static bool TryReadOperand(byte operation, ref LegacyWord69Bytes data, out LegacyWord69Bytes operand) {
            operand = default;
            int length = operation switch {
                4 or 5 or 6 or 7 or 8 or 9 or 10 or 11 or 13 or 14 or 24 or 25 or 29 or 37 or 44 or 50 or 51
                    or 65 or 66 or 67 or 71 or 75 or >= 85 and <= 92 or 94 or 98 or 100 or 102
                    or 104 or 117 or 118 or 185 or 186 => 1,
                2 or 16 or 17 or 18 or 19 or 21 or 22 or >= 26 and <= 28 or >= 30 and <= 36 or >= 38 and <= 43
                    or >= 45 and <= 49 or 69 or 72 or 80 or 93 or 96 or 97 or 99 or 101 or 107 or 109 or 110
                    or 182 or 183 or 184 or 189 or 195 or 197 or 198 => 2,
                73 or 95 => 3,
                20 or 70 or 192 or 194 or 196 or 200 => 4,
                193 or 199 => 5,
                187 => 12,
                83 => 0,
                3 or 12 or 15 or 23 or 68 or 74 or 81 or 82 or 103 or 105 or 106 or 108 or 191 => -1,
                188 or 190 => -2,
                _ => -3
            };

            if (length == -1 && !data.IsEmpty) {
                length = data[0] + (operation == 3 ? 3 : operation == 191 ? 1 : 0);
                // A tab-change marker with length 255 has an internal two-array layout.
                if (operation == 23 && length == 255) return false;
                data = data.Slice(1);
            } else if (length == -2 && data.Length >= 2) {
                length = LegacyWord69Reader.U16(data, 0) - 1;
                data = data.Slice(2);
            }

            if (length < 0 || length > data.Length) return false;
            operand = data.Slice(0, length);
            data = data.Slice(length);
            return true;
        }

        private void AppendTableDefinition(List<byte> output, LegacyWord69Bytes operand) {
            const int oldCellSize = 10;
            const int modernCellSize = 20;
            if (operand.IsEmpty || operand[0] is 0 or > 63) {
                HasUnsupportedFormatting = true;
                return;
            }

            int cells = operand[0];
            int edgesEnd = 1 + 2 * (cells + 1);
            if (edgesEnd > operand.Length || (operand.Length - edgesEnd) % oldCellSize != 0
                || (operand.Length - edgesEnd) / oldCellSize > cells) {
                HasUnsupportedFormatting = true;
                return;
            }

            int descriptors = (operand.Length - edgesEnd) / oldCellSize;
            AppendCode(output, 0xD608);
            AppendWord(output, (ushort)(edgesEnd + descriptors * modernCellSize + 1));
            for (int i = 0; i < edgesEnd; i++) output.Add(operand[i]);
            for (int i = 0; i < descriptors; i++) {
                var descriptor = operand.Slice(edgesEnd + i * oldCellSize, oldCellSize);
                ushort oldFlags = LegacyWord69Reader.U16(descriptor, 0);
                if ((oldFlags & ~3) != 0 || !IsZero(descriptor.Slice(2))) HasUnsupportedFormatting = true;
                AppendWord(output, (ushort)(oldFlags & 3));
                for (int j = 2; j < modernCellSize; j++) output.Add(0);
            }
        }

        internal static IReadOnlyList<string> ReadFonts(LegacyWord69Bytes data, out IReadOnlyList<byte> charsets) {
            var fontCharsets = new List<byte>();
            charsets = fontCharsets;
            var fonts = new List<string>();
            if (data.IsEmpty) return fonts;
            if (data.Length < 2 || LegacyWord69Reader.U16(data, 0) != data.Length) {
                throw new InvalidDataException("The Word 6/95 font table has an invalid length.");
            }

            int offset = 2;
            while (offset < data.Length) {
                int length = data[offset] + 1;
                if (length < 7 || length > data.Length - offset) {
                    throw new InvalidDataException("A Word 6/95 font record extends beyond the table.");
                }

                var name = data.Slice(offset + 6, length - 6);
                int terminator = name.IndexOf((byte)0);
                if (terminator < 0) throw new InvalidDataException("A Word 6/95 font name is not terminated.");
                fonts.Add(DecodeName(name.Slice(0, terminator)));
                fontCharsets.Add(data[offset + 4]);
                offset += length;
            }

            return fonts;
        }

        internal static string DecodeName(LegacyWord69Bytes data) {
            var chars = new char[data.Length];
            for (int i = 0; i < data.Length; i++) chars[i] = LegacyDocPieceTable.DecodeWindows1252(data[i]);
            return new string(chars);
        }

        private static bool IsZero(LegacyWord69Bytes data) {
            for (int i = 0; i < data.Length; i++) {
                if (data[i] != 0) return false;
            }

            return true;
        }

        private static void AppendCode(List<byte> bytes, ushort value) => AppendWord(bytes, value);

        private static void AppendWord(List<byte> bytes, ushort value) {
            bytes.Add((byte)value);
            bytes.Add((byte)(value >> 8));
        }
    }
}
