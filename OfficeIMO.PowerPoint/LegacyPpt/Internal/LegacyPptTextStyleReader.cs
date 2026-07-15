using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Decodes the base StyleTextPropAtom paragraph and character run arrays.</summary>
    internal static class LegacyPptTextStyleReader {
        private const uint CharacterStyleMask = 0x00003EB7U;
        private const uint CharacterTypefaceMask = 1U << 16;
        private const uint CharacterSizeMask = 1U << 17;
        private const uint CharacterColorMask = 1U << 18;
        private const uint CharacterPositionMask = 1U << 19;
        private const uint CharacterOldEastAsianTypefaceMask = 1U << 21;
        private const uint CharacterAnsiTypefaceMask = 1U << 22;
        private const uint CharacterSymbolTypefaceMask = 1U << 23;
        private const uint CharacterUnsupportedExtensionMask = (1U << 20) | (7U << 24);
        private const uint CharacterReservedMask = 0xF800C148U;

        internal static LegacyPptTextBody Read(string text, int rawCharacterCount,
            LegacyPptRecord? styleRecord, LegacyPptColorScheme? colorScheme) {
            if (styleRecord == null) return LegacyPptTextBody.Plain(text);
            try {
                var cursor = new Cursor(styleRecord);
                int styledCharacterCount = checked(rawCharacterCount + 1);
                bool hasParagraphFormatting = ReadParagraphRuns(cursor, styledCharacterCount);
                IReadOnlyList<LegacyPptCharacterRun> runs = ReadCharacterRuns(cursor, text,
                    styledCharacterCount, colorScheme, out bool hasUnprojectedFormatting);
                if (!cursor.IsAtEnd) {
                    throw new InvalidDataException("StyleTextPropAtom contains trailing bytes after its character runs.");
                }
                return new LegacyPptTextBody(text, runs, hasStyleRecord: true,
                    hasParagraphFormatting, hasUnprojectedFormatting);
            } catch (Exception exception) when (exception is InvalidDataException
                                                || exception is OverflowException
                                                || exception is ArgumentOutOfRangeException) {
                return new LegacyPptTextBody(text, Array.Empty<LegacyPptCharacterRun>(),
                    hasStyleRecord: true, hasParagraphFormatting: false,
                    hasUnprojectedCharacterFormatting: true, isStyleTruncated: true);
            }
        }

        private static bool ReadParagraphRuns(Cursor cursor, int characterCount) {
            long covered = 0;
            bool hasFormatting = false;
            while (covered < characterCount) {
                uint count = cursor.ReadUInt32();
                if (count == 0) throw new InvalidDataException("A TextPFRun has a zero character count.");
                covered = checked(covered + count);
                ushort indentLevel = cursor.ReadUInt16();
                uint masks = cursor.ReadUInt32();
                hasFormatting |= indentLevel != 0 || masks != 0;
                SkipParagraphException(cursor, masks);
            }
            ValidateCoverage(covered, characterCount, "paragraph");
            return hasFormatting;
        }

        private static void SkipParagraphException(Cursor cursor, uint masks) {
            if ((masks & 0x0000000FU) != 0) cursor.Skip(2); // bullet flags
            if ((masks & (1U << 7)) != 0) cursor.Skip(2);  // bullet character
            if ((masks & (1U << 4)) != 0) cursor.Skip(2);  // bullet font
            if ((masks & (1U << 6)) != 0) cursor.Skip(2);  // bullet size
            if ((masks & (1U << 5)) != 0) cursor.Skip(4);  // bullet color
            if ((masks & (1U << 11)) != 0) cursor.Skip(2); // alignment
            if ((masks & (1U << 12)) != 0) cursor.Skip(2); // line spacing
            if ((masks & (1U << 13)) != 0) cursor.Skip(2); // space before
            if ((masks & (1U << 14)) != 0) cursor.Skip(2); // space after
            if ((masks & (1U << 8)) != 0) cursor.Skip(2);  // left margin
            if ((masks & (1U << 10)) != 0) cursor.Skip(2); // indent
            if ((masks & (1U << 15)) != 0) cursor.Skip(2); // default tab size
            if ((masks & (1U << 20)) != 0) {
                ushort tabCount = cursor.ReadUInt16();
                cursor.Skip(checked(tabCount * 4));
            }
            if ((masks & (1U << 16)) != 0) cursor.Skip(2); // font alignment
            if ((masks & (7U << 17)) != 0) cursor.Skip(2); // wrap flags
            if ((masks & (1U << 21)) != 0) cursor.Skip(2); // text direction
            if ((masks & 0xFC400200U) != 0) {
                throw new InvalidDataException("A TextPFException uses reserved or extension mask bits.");
            }
            if ((masks & (1U << 23)) != 0) cursor.Skip(2); // bullet blip reference
            if ((masks & (1U << 24)) != 0) cursor.Skip(2); // bullet auto-number scheme
            if ((masks & (1U << 25)) != 0) cursor.Skip(2); // bullet scheme flag
        }

        private static IReadOnlyList<LegacyPptCharacterRun> ReadCharacterRuns(Cursor cursor, string text,
            int characterCount, LegacyPptColorScheme? colorScheme, out bool hasUnprojectedFormatting) {
            var runs = new List<LegacyPptCharacterRun>();
            hasUnprojectedFormatting = false;
            long covered = 0;
            while (covered < characterCount) {
                uint count = cursor.ReadUInt32();
                if (count == 0) throw new InvalidDataException("A TextCFRun has a zero character count.");
                int rawStart = checked((int)covered);
                covered = checked(covered + count);
                uint masks = cursor.ReadUInt32();
                if ((masks & CharacterReservedMask) != 0) {
                    throw new InvalidDataException("A TextCFException uses reserved mask bits.");
                }
                if ((masks & CharacterUnsupportedExtensionMask) != 0) {
                    throw new InvalidDataException("A TextCFException uses an extension record that is not decoded yet.");
                }

                ushort? style = null;
                if ((masks & CharacterStyleMask) != 0) style = cursor.ReadUInt16();
                bool? bold = ReadStyleFlag(masks, style, 0);
                bool? italic = ReadStyleFlag(masks, style, 1);
                bool? underline = ReadStyleFlag(masks, style, 2);
                bool? shadow = ReadStyleFlag(masks, style, 4);
                bool? farEastHint = ReadStyleFlag(masks, style, 5);
                bool? kumi = ReadStyleFlag(masks, style, 7);
                bool? emboss = ReadStyleFlag(masks, style, 9);

                ushort? fontIndex = ReadOptionalUInt16(cursor, masks, CharacterTypefaceMask);
                ushort? oldEastAsianFontIndex = ReadOptionalUInt16(cursor, masks,
                    CharacterOldEastAsianTypefaceMask);
                ushort? ansiFontIndex = ReadOptionalUInt16(cursor, masks, CharacterAnsiTypefaceMask);
                ushort? symbolFontIndex = ReadOptionalUInt16(cursor, masks, CharacterSymbolTypefaceMask);
                short? fontSize = (masks & CharacterSizeMask) != 0 ? cursor.ReadInt16() : null;
                string? color = null;
                byte? schemeIndex = null;
                bool unresolvedColor = false;
                if ((masks & CharacterColorMask) != 0) {
                    byte red = cursor.ReadByte();
                    byte green = cursor.ReadByte();
                    byte blue = cursor.ReadByte();
                    byte index = cursor.ReadByte();
                    if (index == 0xFE) {
                        color = string.Concat(red.ToString("X2"), green.ToString("X2"), blue.ToString("X2"));
                    } else if (index <= 7) {
                        schemeIndex = index;
                        unresolvedColor = colorScheme == null || !colorScheme.TryGetColor(index, out color);
                    } else {
                        unresolvedColor = true;
                    }
                }
                short? position = (masks & CharacterPositionMask) != 0 ? cursor.ReadInt16() : null;

                bool unprojected = shadow.HasValue || farEastHint.HasValue || kumi.HasValue || emboss.HasValue
                    || fontIndex.HasValue || oldEastAsianFontIndex.HasValue || ansiFontIndex.HasValue
                    || symbolFontIndex.HasValue || unresolvedColor;
                hasUnprojectedFormatting |= unprojected;

                int clippedStart = Math.Min(rawStart, text.Length);
                int clippedEnd = Math.Min(checked(rawStart + checked((int)count)), text.Length);
                if (clippedEnd > clippedStart) {
                    runs.Add(new LegacyPptCharacterRun(clippedStart, clippedEnd - clippedStart,
                        text.Substring(clippedStart, clippedEnd - clippedStart), bold, italic, underline,
                        shadow, farEastHint, kumi, emboss, fontIndex, oldEastAsianFontIndex,
                        ansiFontIndex, symbolFontIndex, fontSize, color, schemeIndex, position, unprojected));
                }
            }
            ValidateCoverage(covered, characterCount, "character");
            return runs;
        }

        private static ushort? ReadOptionalUInt16(Cursor cursor, uint masks, uint mask) =>
            (masks & mask) != 0 ? cursor.ReadUInt16() : null;

        private static bool? ReadStyleFlag(uint masks, ushort? style, int bit) =>
            (masks & (1U << bit)) == 0 ? null : (style.GetValueOrDefault() & (1U << bit)) != 0;

        private static void ValidateCoverage(long covered, int expected, string kind) {
            if (covered != expected) {
                throw new InvalidDataException(
                    $"StyleTextPropAtom {kind} runs cover {covered} characters instead of {expected}.");
            }
        }

        private sealed class Cursor {
            private readonly LegacyPptRecord _record;

            internal Cursor(LegacyPptRecord record) {
                _record = record ?? throw new ArgumentNullException(nameof(record));
            }

            internal int Offset { get; private set; }

            internal bool IsAtEnd => Offset == _record.PayloadLength;

            internal byte ReadByte() {
                byte value = _record.ReadByte(Offset);
                Offset++;
                return value;
            }

            internal ushort ReadUInt16() {
                ushort value = _record.ReadUInt16(Offset);
                Offset = checked(Offset + 2);
                return value;
            }

            internal short ReadInt16() => unchecked((short)ReadUInt16());

            internal uint ReadUInt32() {
                uint value = _record.ReadUInt32(Offset);
                Offset = checked(Offset + 4);
                return value;
            }

            internal void Skip(int count) {
                if (count < 0 || Offset > _record.PayloadLength - count) {
                    throw new InvalidDataException("StyleTextPropAtom is truncated.");
                }
                Offset = checked(Offset + count);
            }
        }
    }
}
