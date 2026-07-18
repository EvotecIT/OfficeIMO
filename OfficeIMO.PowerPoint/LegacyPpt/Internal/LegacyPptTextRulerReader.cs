using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Decodes TextRulerAtom tab, margin, and indentation settings.</summary>
    internal static class LegacyPptTextRulerReader {
        private const uint KnownMask = 0x00001FFFU;

        internal static LegacyPptTextRuler? Read(LegacyPptRecord? record, out bool isTruncated) {
            isTruncated = false;
            if (record == null) return null;
            try {
                var cursor = new LegacyPptTextPropertyCursor(record, "TextRulerAtom");
                uint mask = cursor.ReadUInt32();
                if ((mask & ~KnownMask) != 0) {
                    throw new InvalidDataException("TextRulerAtom uses reserved mask bits.");
                }
                short? levelCount = (mask & (1U << 1)) != 0 ? cursor.ReadInt16() : null;
                if (levelCount.HasValue && (levelCount.Value < 0 || levelCount.Value > 5)) {
                    throw new InvalidDataException("TextRulerAtom has a level count outside the zero-to-five range.");
                }
                short? defaultTabSize = (mask & 1U) != 0 ? cursor.ReadInt16() : null;
                bool hasUnprojectedFormatting = defaultTabSize < 0;
                IReadOnlyList<LegacyPptTabStop> tabStops = (mask & (1U << 2)) != 0
                    ? LegacyPptTextPropertyReader.ReadTabStops(cursor)
                    : Array.Empty<LegacyPptTabStop>();
                var levels = new List<LegacyPptTextRulerLevel>();
                for (ushort level = 0; level < 5; level++) {
                    short? leftMargin = (mask & (1U << (3 + level))) != 0
                        ? cursor.ReadInt16()
                        : null;
                    short? indent = (mask & (1U << (8 + level))) != 0
                        ? cursor.ReadInt16()
                        : null;
                    if (leftMargin.HasValue || indent.HasValue) {
                        levels.Add(new LegacyPptTextRulerLevel(level, leftMargin, indent));
                    }
                }
                if (!cursor.IsAtEnd) {
                    throw new InvalidDataException("TextRulerAtom contains trailing bytes.");
                }
                return new LegacyPptTextRuler(levelCount, defaultTabSize, tabStops, levels,
                    hasUnprojectedFormatting);
            } catch (Exception exception) when (exception is InvalidDataException
                                                || exception is OverflowException
                                                || exception is ArgumentOutOfRangeException) {
                isTruncated = true;
                return null;
            }
        }
    }
}
