using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Decodes reusable binary PowerPoint paragraph and character exceptions.</summary>
    internal static class LegacyPptTextPropertyReader {
        private const int MaximumTabStopCount = 4096;
        private const uint CharacterStyleMask = 0x00003EB7U;
        private const uint CharacterTypefaceMask = 1U << 16;
        private const uint CharacterSizeMask = 1U << 17;
        private const uint CharacterColorMask = 1U << 18;
        private const uint CharacterPositionMask = 1U << 19;
        private const uint CharacterOldEastAsianTypefaceMask = 1U << 21;
        private const uint CharacterAnsiTypefaceMask = 1U << 22;
        private const uint CharacterSymbolTypefaceMask = 1U << 23;
        private const uint CharacterUnsupportedExtensionMask = (1U << 20) | (7U << 24);
        private const uint CharacterReservedMask = 0xF8000000U;

        internal static LegacyPptParagraphRun ReadParagraphException(
            LegacyPptTextPropertyCursor cursor, int start, int length, ushort indentLevel,
            LegacyPptColorScheme? colorScheme, IReadOnlyDictionary<ushort, LegacyPptFont>? fonts,
            bool allowRulerFields, out bool hasUnprojectedFormatting) {
            if (indentLevel > 4) {
                throw new InvalidDataException("A TextPFException indent level is greater than four.");
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

            short? rawLeftMargin = (masks & (1U << 8)) != 0 ? cursor.ReadInt16() : null;
            short? rawIndent = (masks & (1U << 10)) != 0 ? cursor.ReadInt16() : null;
            short? rawDefaultTabSize = (masks & (1U << 15)) != 0 ? cursor.ReadInt16() : null;
            IReadOnlyList<LegacyPptTabStop> rawTabStops = (masks & (1U << 20)) != 0
                ? ReadTabStops(cursor)
                : Array.Empty<LegacyPptTabStop>();

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

            bool unresolvedBulletFontValue = false;
            string? bulletTypeface = bulletHasFont == true
                ? ResolveFont(bulletFontIndex, fonts, out unresolvedBulletFontValue)
                : null;
            bool unresolvedBulletFont = bulletHasFont == true && unresolvedBulletFontValue;
            bool bulletSizeUnprojected = bulletHasSize == true && bulletSize.HasValue
                && !((bulletSize.Value >= 25 && bulletSize.Value <= 400)
                    || (bulletSize.Value >= -4000 && bulletSize.Value <= -1));
            bool hasRulerOnlyField = !allowRulerFields && (rawLeftMargin.HasValue
                || rawIndent.HasValue || rawDefaultTabSize.HasValue || rawTabStops.Count != 0
                || (masks & (1U << 20)) != 0);
            bool invalidDefaultTab = allowRulerFields && rawDefaultTabSize < 0;
            hasUnprojectedFormatting = hasRulerOnlyField || unresolvedBulletFont
                || (bulletHasColor == true && unresolvedBulletColor)
                || bulletSizeUnprojected || invalidDefaultTab;

            return new LegacyPptParagraphRun(start, length, indentLevel, hasBullet,
                bulletHasFont, bulletHasColor, bulletHasSize, bulletCharacter, bulletFontIndex,
                bulletTypeface, bulletSize, bulletColor, bulletColorSchemeIndex, alignment,
                lineSpacing, spaceBefore, spaceAfter, fontAlignment, characterWrap, wordWrap,
                overflow, textDirection, hasUnprojectedFormatting,
                allowRulerFields ? rawLeftMargin : null,
                allowRulerFields ? rawIndent : null,
                allowRulerFields && !invalidDefaultTab ? rawDefaultTabSize : null,
                allowRulerFields ? rawTabStops : null);
        }

        internal static LegacyPptCharacterRun ReadCharacterException(
            LegacyPptTextPropertyCursor cursor, int start, int length, string text,
            LegacyPptColorScheme? colorScheme, IReadOnlyDictionary<ushort, LegacyPptFont>? fonts,
            out bool hasUnprojectedFormatting) {
            uint masks = cursor.ReadUInt32();
            if ((masks & CharacterReservedMask) != 0) {
                throw new InvalidDataException(
                    $"A TextCFException uses reserved mask bits in mask 0x{masks:X8} at payload offset {cursor.Offset - 4}.");
            }
            if ((masks & CharacterUnsupportedExtensionMask) != 0) {
                throw new InvalidDataException("A TextCFException uses an extension record that is not decoded yet.");
            }

            ushort? style = (masks & CharacterStyleMask) != 0 ? cursor.ReadUInt16() : null;
            byte? ppt9RunId = (masks & 0x00003C00U) != 0
                ? checked((byte)((style.GetValueOrDefault() >> 10) & 0x0F))
                : null;
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
            hasUnprojectedFormatting = shadow == true || farEastHint == true || emboss == true
                || unresolvedPrimaryFont || unresolvedEastAsianFont
                || unresolvedAnsiFont || unresolvedSymbolFont || unresolvedColor;

            return new LegacyPptCharacterRun(start, length, text, bold, italic, underline,
                shadow, farEastHint, kumi, emboss, fontIndex, oldEastAsianFontIndex,
                ansiFontIndex, symbolFontIndex, typeface, oldEastAsianTypeface, ansiTypeface,
                symbolTypeface, fontSize, color, schemeIndex, position,
                hasUnprojectedFormatting, ppt9RunId);
        }

        internal static IReadOnlyList<LegacyPptTabStop> ReadTabStops(
            LegacyPptTextPropertyCursor cursor) {
            ushort count = cursor.ReadUInt16();
            if (count > MaximumTabStopCount) {
                throw new InvalidDataException(
                    $"A binary PowerPoint tab-stop list cannot contain more than {MaximumTabStopCount} entries.");
            }
            var tabStops = new List<LegacyPptTabStop>(count);
            for (int index = 0; index < count; index++) {
                short position = cursor.ReadInt16();
                ushort alignment = cursor.ReadUInt16();
                if (alignment > 3) {
                    throw new InvalidDataException("A binary PowerPoint tab-stop alignment is invalid.");
                }
                tabStops.Add(new LegacyPptTabStop(position, (LegacyPptTabAlignment)alignment));
            }
            return tabStops;
        }

        private static ushort? ReadOptionalUInt16(LegacyPptTextPropertyCursor cursor,
            uint masks, uint mask) => (masks & mask) != 0 ? cursor.ReadUInt16() : null;

        private static short? ReadSpacing(LegacyPptTextPropertyCursor cursor, uint masks, int bit) {
            if ((masks & (1U << bit)) == 0) return null;
            short value = cursor.ReadInt16();
            if (value > 13200) {
                throw new InvalidDataException("A TextPFException spacing value is greater than 13200 percent.");
            }
            return value;
        }

        private static string? ReadColor(LegacyPptTextPropertyCursor cursor,
            LegacyPptColorScheme? colorScheme, out byte? schemeIndex, out bool unresolved) {
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
            if (index == 0xFF) return null;
            unresolved = true;
            return null;
        }

        private static string? ResolveFont(ushort? index,
            IReadOnlyDictionary<ushort, LegacyPptFont>? fonts, out bool unresolved) {
            if (index == ushort.MaxValue) {
                unresolved = false;
                return null;
            }
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
    }
}
