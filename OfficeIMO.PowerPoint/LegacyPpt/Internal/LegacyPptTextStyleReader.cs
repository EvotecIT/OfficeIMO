using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Decodes the base StyleTextPropAtom paragraph and character run arrays.</summary>
    internal static class LegacyPptTextStyleReader {
        internal static LegacyPptTextBody Read(string text, int rawCharacterCount,
            LegacyPptRecord? styleRecord, LegacyPptColorScheme? colorScheme,
            IReadOnlyDictionary<ushort, LegacyPptFont>? fonts = null,
            LegacyPptTextType? textType = null, LegacyPptTextRuler? ruler = null,
            bool hasRulerRecord = false, bool isRulerTruncated = false) {
            if (styleRecord == null) {
                return new LegacyPptTextBody(text, Array.Empty<LegacyPptCharacterRun>(),
                    Array.Empty<LegacyPptParagraphRun>(), hasStyleRecord: false,
                    hasUnprojectedCharacterFormatting: false,
                    hasUnprojectedParagraphFormatting: ruler?.HasUnprojectedFormatting == true,
                    textType: textType, ruler: ruler, hasRulerRecord: hasRulerRecord,
                    isRulerTruncated: isRulerTruncated);
            }
            try {
                var cursor = new LegacyPptTextPropertyCursor(styleRecord, "StyleTextPropAtom");
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
                    hasUnprojectedParagraphFormatting: hasUnprojectedParagraphFormatting
                        || ruler?.HasUnprojectedFormatting == true,
                    textType: textType, ruler: ruler, hasRulerRecord: hasRulerRecord,
                    isRulerTruncated: isRulerTruncated);
            } catch (Exception exception) when (exception is InvalidDataException
                                                || exception is OverflowException
                                                || exception is ArgumentOutOfRangeException) {
                return new LegacyPptTextBody(text, Array.Empty<LegacyPptCharacterRun>(),
                    Array.Empty<LegacyPptParagraphRun>(), hasStyleRecord: true,
                    hasUnprojectedCharacterFormatting: true,
                    hasUnprojectedParagraphFormatting: true, isStyleTruncated: true,
                    textType: textType, ruler: ruler, hasRulerRecord: hasRulerRecord,
                    isRulerTruncated: isRulerTruncated);
            }
        }

        private static IReadOnlyList<LegacyPptParagraphRun> ReadParagraphRuns(
            LegacyPptTextPropertyCursor cursor, string text, int characterCount,
            LegacyPptColorScheme? colorScheme, IReadOnlyDictionary<ushort, LegacyPptFont>? fonts,
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
                int clippedStart = Math.Min(rawStart, text.Length);
                int clippedEnd = Math.Min(checked(rawStart + checked((int)count)), text.Length);
                LegacyPptParagraphRun decoded = LegacyPptTextPropertyReader.ReadParagraphException(
                    cursor, clippedStart, Math.Max(0, clippedEnd - clippedStart), indentLevel,
                    colorScheme, fonts, allowRulerFields: false, out bool unprojected);
                hasUnprojectedFormatting |= unprojected;
                if (clippedEnd > clippedStart) runs.Add(decoded);
            }
            ValidateCoverage(covered, characterCount, "paragraph");
            return runs;
        }

        private static IReadOnlyList<LegacyPptCharacterRun> ReadCharacterRuns(
            LegacyPptTextPropertyCursor cursor, string text, int characterCount,
            LegacyPptColorScheme? colorScheme, IReadOnlyDictionary<ushort, LegacyPptFont>? fonts,
            out bool hasUnprojectedFormatting) {
            var runs = new List<LegacyPptCharacterRun>();
            hasUnprojectedFormatting = false;
            long covered = 0;
            while (covered < characterCount) {
                uint count = cursor.ReadUInt32();
                if (count == 0) throw new InvalidDataException("A TextCFRun has a zero character count.");
                int rawStart = checked((int)covered);
                covered = checked(covered + count);
                int clippedStart = Math.Min(rawStart, text.Length);
                int clippedEnd = Math.Min(checked(rawStart + checked((int)count)), text.Length);
                string runText = clippedEnd > clippedStart
                    ? text.Substring(clippedStart, clippedEnd - clippedStart)
                    : string.Empty;
                LegacyPptCharacterRun decoded = LegacyPptTextPropertyReader.ReadCharacterException(
                    cursor, clippedStart, Math.Max(0, clippedEnd - clippedStart), runText,
                    colorScheme, fonts, out bool unprojected);
                hasUnprojectedFormatting |= unprojected;
                if (clippedEnd > clippedStart) runs.Add(decoded);
            }
            ValidateCoverage(covered, characterCount, "character");
            return runs;
        }

        private static void ValidateCoverage(long covered, int expected, string kind) {
            if (covered != expected) {
                throw new InvalidDataException(
                    $"StyleTextPropAtom {kind} runs cover {covered} characters instead of {expected}.");
            }
        }
    }
}
