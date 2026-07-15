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
            LegacyPptRecord? styleRecord, LegacyPptColorScheme? colorScheme,
            IReadOnlyDictionary<ushort, LegacyPptFont>? fonts = null) {
            if (styleRecord == null) return LegacyPptTextBody.Plain(text);
            try {
                var cursor = new Cursor(styleRecord);
                int styledCharacterCount = checked(rawCharacterCount + 1);
                IReadOnlyList<LegacyPptParagraphRun> paragraphRuns = ReadParagraphRuns(cursor, text,
                    styledCharacterCount, colorScheme, fonts,
                    out bool hasUnprojectedParagraphFormatting);
                IReadOnlyList<LegacyPptCharacterRun> characterRuns = ReadCharacterRuns(cursor, text,
                    styledCharacterCount, colorScheme, fonts,
                    out bool hasUnprojectedCharacterFormatting);
                if (!cursor.IsAtEnd) {
                    throw new InvalidDataException("StyleTextPropAtom contains trailing bytes after its character runs.");
                }
                return new LegacyPptTextBody(text, characterRuns, paragraphRuns,
                    hasStyleRecord: true, hasUnprojectedCharacterFormatting,
                    hasUnprojectedParagraphFormatting);
            } catch (Exception exception) when (exception is InvalidDataException
                                                || exception is OverflowException
                                                || exception is ArgumentOutOfRangeException) {
                return new LegacyPptTextBody(text, Array.Empty<LegacyPptCharacterRun>(),
                    Array.Empty<LegacyPptParagraphRun>(), hasStyleRecord: true,
                    hasUnprojectedCharacterFormatting: true,
                    hasUnprojectedParagraphFormatting: true, isStyleTruncated: true);
            }
        }

        private static IReadOnlyList<LegacyPptParagraphRun> ReadParagraphRuns(Cursor cursor,
            string text, int characterCount, LegacyPptColorScheme? colorScheme,
            IReadOnlyDictionary<ushort, LegacyPptFont>? fonts,
            out bool hasUnprojectedFormatting) {
            var runs = new List<LegacyPptParagraphRun>();
            hasUnprojectedFormatting = false;
            long covered = 0;
            while (covered < characterCount) {
                uint count = cursor.ReadUInt32();
                if (count == 0) throw new InvalidDataException("A TextPFRun has a zero character count.");
                int rawStart = checked((int)covered);
                covered = checked(covered + count);
                ushort indentLevel = cursor.ReadUInt16();
                if (indentLevel > 4) {
                    throw new InvalidDataException("A TextPFRun indent level is greater than four.");
                }
                uint masks = cursor.ReadUInt32();
                if ((masks & 0x03800000U) != 0) {
                    throw new InvalidDataException("A TextPFException uses an extension field that is not decoded yet.");
                }

                ushort? bulletFlags = (masks & 0x0000000FU) != 0 ? cursor.ReadUInt16() : null;
                bool? hasBullet = ReadMaskedFlag(masks, bulletFlags, 0);
                bool? bulletHasFont = ReadMaskedFlag(masks, bulletFlags, 1);
                bool? bulletHasColor = ReadMaskedFlag(masks, bulletFlags, 2);
                bool? bulletHasSize = ReadMaskedFlag(masks, bulletFlags, 3);
                char? bulletCharacter = null;
                if ((masks & (1U << 7)) != 0) {
                    ushort character = cursor.ReadUInt16();
                    if (character == 0) throw new InvalidDataException("A TextPFException bullet character is NUL.");
                    bulletCharacter = (char)character;
                }
                ushort? bulletFontIndex = ReadOptionalUInt16(cursor, masks, 1U << 4);
                short? bulletSize = (masks & (1U << 6)) != 0 ? cursor.ReadInt16() : null;
                string? bulletColor = null;
                byte? bulletColorSchemeIndex = null;
                bool unresolvedBulletColor = false;
                if ((masks & (1U << 5)) != 0) {
                    bulletColor = ReadColor(cursor, colorScheme, out bulletColorSchemeIndex,
                        out unresolvedBulletColor);
                }
                LegacyPptTextAlignment? alignment = null;
                if ((masks & (1U << 11)) != 0) {
                    ushort value = cursor.ReadUInt16();
                    if (value > 6) throw new InvalidDataException("A TextPFException alignment is invalid.");
                    alignment = (LegacyPptTextAlignment)value;
                }
                short? lineSpacing = ReadSpacing(cursor, masks, 12);
                short? spaceBefore = ReadSpacing(cursor, masks, 13);
                short? spaceAfter = ReadSpacing(cursor, masks, 14);

                bool hasRulerOnlyField = false;
                if ((masks & (1U << 8)) != 0) {
                    cursor.Skip(2); // left margin is invalid in TextPFRun
                    hasRulerOnlyField = true;
                }
                if ((masks & (1U << 10)) != 0) {
                    cursor.Skip(2); // indent is invalid in TextPFRun
                    hasRulerOnlyField = true;
                }
                if ((masks & (1U << 15)) != 0) {
                    cursor.Skip(2); // default tab size is invalid in TextPFRun
                    hasRulerOnlyField = true;
                }
                if ((masks & (1U << 20)) != 0) {
                    ushort tabCount = cursor.ReadUInt16();
                    cursor.Skip(checked(tabCount * 4));
                    hasRulerOnlyField = true;
                }

                LegacyPptFontAlignment? fontAlignment = null;
                if ((masks & (1U << 16)) != 0) {
                    ushort value = cursor.ReadUInt16();
                    if (value > 3) throw new InvalidDataException("A TextPFException font alignment is invalid.");
                    fontAlignment = (LegacyPptFontAlignment)value;
                }
                ushort? wrapFlags = (masks & (7U << 17)) != 0 ? cursor.ReadUInt16() : null;
                bool? characterWrap = ReadMaskedFlag(masks >> 17, wrapFlags, 0);
                bool? wordWrap = ReadMaskedFlag(masks >> 17, wrapFlags, 1);
                bool? overflow = ReadMaskedFlag(masks >> 17, wrapFlags, 2);
                LegacyPptTextDirection? textDirection = null;
                if ((masks & (1U << 21)) != 0) {
                    ushort value = cursor.ReadUInt16();
                    if (value > 1) throw new InvalidDataException("A TextPFException text direction is invalid.");
                    textDirection = (LegacyPptTextDirection)value;
                }

                string? bulletTypeface = ResolveFont(bulletFontIndex, fonts,
                    out bool unresolvedBulletFont);
                bool bulletSizeUnprojected = bulletSize.HasValue
                    && !((bulletSize.Value >= 25 && bulletSize.Value <= 400)
                        || (bulletSize.Value >= -4000 && bulletSize.Value <= -1));
                bool unprojected = hasRulerOnlyField || unresolvedBulletFont
                    || unresolvedBulletColor || bulletSizeUnprojected;
                hasUnprojectedFormatting |= unprojected;

                int clippedStart = Math.Min(rawStart, text.Length);
                int clippedEnd = Math.Min(checked(rawStart + checked((int)count)), text.Length);
                if (clippedEnd > clippedStart) {
                    runs.Add(new LegacyPptParagraphRun(clippedStart, clippedEnd - clippedStart,
                        indentLevel, hasBullet, bulletHasFont, bulletHasColor, bulletHasSize,
                        bulletCharacter, bulletFontIndex, bulletTypeface, bulletSize, bulletColor,
                        bulletColorSchemeIndex, alignment, lineSpacing, spaceBefore, spaceAfter,
                        fontAlignment, characterWrap, wordWrap, overflow, textDirection, unprojected));
                }
            }
            ValidateCoverage(covered, characterCount, "paragraph");
            return runs;
        }

        private static IReadOnlyList<LegacyPptCharacterRun> ReadCharacterRuns(Cursor cursor, string text,
            int characterCount, LegacyPptColorScheme? colorScheme,
            IReadOnlyDictionary<ushort, LegacyPptFont>? fonts,
            out bool hasUnprojectedFormatting) {
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
                if (fontSize.HasValue && (fontSize.Value < 1 || fontSize.Value > 4000)) {
                    throw new InvalidDataException("A TextCFException font size is outside the 1-to-4000 point range.");
                }
                string? color = null;
                byte? schemeIndex = null;
                bool unresolvedColor = false;
                if ((masks & CharacterColorMask) != 0) {
                    color = ReadColor(cursor, colorScheme, out schemeIndex, out unresolvedColor);
                }
                short? position = (masks & CharacterPositionMask) != 0 ? cursor.ReadInt16() : null;
                if (position.HasValue && (position.Value < -100 || position.Value > 100)) {
                    throw new InvalidDataException("A TextCFException baseline position is outside the -100-to-100 percent range.");
                }

                string? typeface = ResolveFont(fontIndex, fonts, out bool unresolvedPrimaryFont);
                string? oldEastAsianTypeface = ResolveFont(oldEastAsianFontIndex, fonts,
                    out bool unresolvedEastAsianFont);
                string? ansiTypeface = ResolveFont(ansiFontIndex, fonts, out bool unresolvedAnsiFont);
                string? symbolTypeface = ResolveFont(symbolFontIndex, fonts, out bool unresolvedSymbolFont);

                bool unprojected = shadow.HasValue || farEastHint.HasValue || kumi.HasValue || emboss.HasValue
                    || unresolvedPrimaryFont || unresolvedEastAsianFont || unresolvedAnsiFont
                    || unresolvedSymbolFont || unresolvedColor;
                hasUnprojectedFormatting |= unprojected;

                int clippedStart = Math.Min(rawStart, text.Length);
                int clippedEnd = Math.Min(checked(rawStart + checked((int)count)), text.Length);
                if (clippedEnd > clippedStart) {
                    runs.Add(new LegacyPptCharacterRun(clippedStart, clippedEnd - clippedStart,
                        text.Substring(clippedStart, clippedEnd - clippedStart), bold, italic, underline,
                        shadow, farEastHint, kumi, emboss, fontIndex, oldEastAsianFontIndex,
                        ansiFontIndex, symbolFontIndex, typeface, oldEastAsianTypeface, ansiTypeface,
                        symbolTypeface, fontSize, color, schemeIndex, position, unprojected));
                }
            }
            ValidateCoverage(covered, characterCount, "character");
            return runs;
        }

        private static ushort? ReadOptionalUInt16(Cursor cursor, uint masks, uint mask) =>
            (masks & mask) != 0 ? cursor.ReadUInt16() : null;

        private static short? ReadSpacing(Cursor cursor, uint masks, int bit) {
            if ((masks & (1U << bit)) == 0) return null;
            short value = cursor.ReadInt16();
            if (value > 13200) {
                throw new InvalidDataException("A TextPFException spacing value is greater than 13200 percent.");
            }
            return value;
        }

        private static string? ReadColor(Cursor cursor, LegacyPptColorScheme? colorScheme,
            out byte? schemeIndex, out bool unresolved) {
            byte red = cursor.ReadByte();
            byte green = cursor.ReadByte();
            byte blue = cursor.ReadByte();
            byte index = cursor.ReadByte();
            schemeIndex = null;
            unresolved = false;
            if (index == 0xFE) {
                return string.Concat(red.ToString("X2"), green.ToString("X2"), blue.ToString("X2"));
            }
            if (index <= 7) {
                schemeIndex = index;
                if (colorScheme != null && colorScheme.TryGetColor(index, out string? resolved)) {
                    return resolved;
                }
            }
            unresolved = true;
            return null;
        }

        private static string? ResolveFont(ushort? index,
            IReadOnlyDictionary<ushort, LegacyPptFont>? fonts, out bool unresolved) {
            unresolved = index.HasValue && (fonts == null || !fonts.TryGetValue(index.Value, out _));
            return index.HasValue && fonts != null
                && fonts.TryGetValue(index.Value, out LegacyPptFont? font)
                ? font.Typeface
                : null;
        }

        private static bool? ReadStyleFlag(uint masks, ushort? style, int bit) =>
            (masks & (1U << bit)) == 0 ? null : (style.GetValueOrDefault() & (1U << bit)) != 0;

        private static bool? ReadMaskedFlag(uint masks, ushort? values, int bit) =>
            (masks & (1U << bit)) == 0 ? null : (values.GetValueOrDefault() & (1U << bit)) != 0;

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
