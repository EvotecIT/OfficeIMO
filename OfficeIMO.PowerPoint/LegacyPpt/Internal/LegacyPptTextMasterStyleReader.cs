using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Decodes the five-level base TextMasterStyleAtom structure.</summary>
    internal static class LegacyPptTextMasterStyleReader {
        internal static LegacyPptTextMasterStyle? Read(LegacyPptRecord record,
            LegacyPptColorScheme? colorScheme,
            IReadOnlyDictionary<ushort, LegacyPptFont>? fonts) {
            if (record == null) throw new ArgumentNullException(nameof(record));
            if (!TryMapTextType(record.Instance, out LegacyPptTextType textType)) return null;
            try {
                var cursor = new LegacyPptTextPropertyCursor(record, "TextMasterStyleAtom");
                ushort levelCount = cursor.ReadUInt16();
                if (levelCount > 5) {
                    throw new InvalidDataException("TextMasterStyleAtom contains more than five levels.");
                }
                bool hasExplicitLevel = record.Instance >= 5;
                bool hasUnprojectedFormatting = false;
                var levels = new List<LegacyPptTextMasterStyleLevel>(levelCount);
                var seenLevels = new HashSet<ushort>();
                for (ushort index = 0; index < levelCount; index++) {
                    ushort level = hasExplicitLevel ? cursor.ReadUInt16() : index;
                    if (level >= levelCount || level > 4 || !seenLevels.Add(level)) {
                        throw new InvalidDataException("TextMasterStyleAtom contains an invalid or duplicate level index.");
                    }
                    LegacyPptParagraphRun paragraph = LegacyPptTextPropertyReader.ReadParagraphException(
                        cursor, start: 0, length: 0, level, colorScheme, fonts,
                        allowRulerFields: true, out bool paragraphUnprojected);
                    LegacyPptCharacterRun character = LegacyPptTextPropertyReader.ReadCharacterException(
                        cursor, start: 0, length: 0, text: string.Empty, colorScheme, fonts,
                        out bool characterUnprojected);
                    hasUnprojectedFormatting |= paragraphUnprojected || characterUnprojected;
                    levels.Add(new LegacyPptTextMasterStyleLevel(level, paragraph, character));
                }
                if (!cursor.IsAtEnd) {
                    throw new InvalidDataException("TextMasterStyleAtom contains trailing bytes.");
                }
                return new LegacyPptTextMasterStyle(textType, levels.OrderBy(level => level.Level).ToArray(),
                    hasUnprojectedFormatting, isTruncated: false);
            } catch (Exception exception) when (exception is InvalidDataException
                                                || exception is OverflowException
                                                || exception is ArgumentOutOfRangeException) {
                return new LegacyPptTextMasterStyle(textType,
                    Array.Empty<LegacyPptTextMasterStyleLevel>(),
                    hasUnprojectedFormatting: true, isTruncated: true);
            }
        }

        private static bool TryMapTextType(ushort value, out LegacyPptTextType textType) {
            if (value == 0 || value == 1 || value == 2 || value == 4
                || value == 5 || value == 6 || value == 7 || value == 8) {
                textType = (LegacyPptTextType)value;
                return true;
            }
            textType = default;
            return false;
        }
    }
}
