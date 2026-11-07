using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Decodes Word 6/95 section records into the shared editable page layout model.</summary>
    internal static class LegacyWord69Sections {
        internal static IReadOnlyList<LegacyDocSection> Read(
            LegacyWord69Bytes bytes, int usedEnd, LegacyWord69Bytes directory, int bodyLength, int textLength,
            out IReadOnlyList<byte> headerMasks, out bool unsupported) {
            var masks = new List<byte>();
            headerMasks = masks;
            unsupported = false;
            if (directory.IsEmpty) {
                masks.Add(0);
                return new[] { new LegacyDocSection(0, bodyLength, ReadProperties(default, out _, out unsupported)) };
            }
            if (directory.Length < 20 || (directory.Length - 4) % 16 != 0) {
                throw new InvalidDataException("The Word 6/95 section directory has an invalid length.");
            }

            int count = (directory.Length - 4) / 16;
            var sections = new List<LegacyDocSection>(count);
            int previousEnd = 0;
            for (int index = 0; index < count; index++) {
                int start = LegacyWord69Reader.I32(directory, index * 4);
                int end = LegacyWord69Reader.I32(directory, (index + 1) * 4);
                if (start != previousEnd || end <= start || end > textLength || start >= bodyLength) {
                    throw new InvalidDataException("The Word 6/95 section directory contains an invalid text range.");
                }

                int descriptor = (count + 1) * 4 + index * 12;
                int location = LegacyWord69Reader.I32(directory, descriptor + 2);
                LegacyWord69Bytes properties = default;
                if (location != -1) {
                    if (location < 0x192 || location > usedEnd - 2) {
                        throw new InvalidDataException("A Word 6/95 section property block points outside the used stream.");
                    }
                    int length = LegacyWord69Reader.U16(bytes, location);
                    if (length > usedEnd - location - 2) {
                        throw new InvalidDataException("A Word 6/95 section property block extends beyond the used stream.");
                    }
                    properties = bytes.Slice(location + 2, length);
                }

                var format = ReadProperties(properties, out byte mask, out bool sectionUnsupported);
                sections.Add(new LegacyDocSection(start, Math.Min(end, bodyLength), format));
                masks.Add(mask);
                unsupported |= sectionUnsupported;
                previousEnd = end;
            }
            if (previousEnd < bodyLength) {
                throw new InvalidDataException("The Word 6/95 section directory does not cover the body text.");
            }
            return sections;
        }

        private static LegacyDocSectionFormat ReadProperties(
            LegacyWord69Bytes data, out byte headerMask, out bool unsupported) {
            headerMask = 0;
            unsupported = false;
            EndnotePositionValues? endnotePosition = null;
            // The binary format's standard SEP uses letter paper, 1.25-inch side margins,
            // 1-inch top/bottom margins and half-inch header/footer distances.
            var output = new List<byte>();
            Append(0xB01F, 12240);
            Append(0xB020, 15840);
            Append(0xB021, 1800);
            Append(0xB022, 1800);
            Append(0x9023, 1440);
            Append(0x9024, 1440);
            Append(0xB017, 720);
            Append(0xB018, 720);
            while (!data.IsEmpty) {
                byte operation = data[0];
                data = data.Slice(1);
                if (operation == 0) continue;
                int length = OperandLength(operation);
                if (length < 0 || length > data.Length) {
                    unsupported = true;
                    break;
                }
                var operand = data.Slice(0, length);
                data = data.Slice(length);
                if (operation == 153) {
                    headerMask = operand[0];
                    if ((headerMask & 0xC0) != 0) unsupported = true;
                    continue;
                }
                if (operation == 151) {
                    endnotePosition = operand[0] == 0 ? EndnotePositionValues.DocumentEnd : EndnotePositionValues.SectionEnd;
                    continue;
                }
                if (operation is 168 or 169 && LegacyWord69Reader.U16(operand, 0) > short.MaxValue) {
                    unsupported = true;
                }
                ushort code = Code(operation);
                if (code == 0) {
                    // Printer tray/paper codes do not change editable page geometry.
                    if (operation is not (140 or 141 or 163 or 171)) unsupported = true;
                    continue;
                }
                output.Add((byte)code);
                output.Add((byte)(code >> 8));
                for (int index = 0; index < operand.Length; index++) output.Add(operand[index]);
            }
            byte[] normalized = output.ToArray();
            var result = LegacyDocSectionFormattingReader.ReadSepxGrpprl(normalized, 0, normalized.Length, out string? warning);
            unsupported |= warning != null;
            return result.WithEndnotePosition(endnotePosition);

            void Append(ushort code, ushort value) {
                output.Add((byte)code);
                output.Add((byte)(code >> 8));
                output.Add((byte)value);
                output.Add((byte)(value >> 8));
            }
        }

        private static int OperandLength(byte operation) => operation switch {
            131 or 132 or 138 or 139 or 142 or 143 or 146 or 147 or 150 or 151 or 152 or 153
                or 158 or 159 or 162 => 1,
            140 or 141 or 144 or 145 or 148 or 149 or 154 or 155 or 156 or 157 or 160 or 161
                or 164 or 165 or 166 or 167 or 168 or 169 or 170 or 171 => 2,
            136 or 137 => 3,
            163 => 0,
            _ => -1
        };

        private static ushort Code(byte operation) => operation switch {
            136 => 0xF203, 137 => 0xF204, 138 => 0x3005,
            142 => 0x3009, 143 => 0x300A, 144 => 0x500B, 145 => 0x900C,
            147 => 0x300E, 150 => 0x3011, 152 => 0x3013, 154 => 0x5015,
            155 => 0x9016, 156 => 0xB017, 157 => 0xB018, 158 => 0x3019,
            159 => 0x301A, 160 => 0x501B, 161 => 0x501C, 162 => 0x301D,
            164 => 0xB01F, 165 => 0xB020, 166 => 0xB021, 167 => 0xB022,
            168 => 0x9023, 169 => 0x9024, 170 => 0xB025,
            _ => 0
        };
    }
}
