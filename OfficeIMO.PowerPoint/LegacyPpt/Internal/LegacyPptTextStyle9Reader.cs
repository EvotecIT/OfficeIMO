using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Decodes and links PPT9 extended text properties to base character runs.</summary>
    internal static class LegacyPptTextStyle9Reader {
        private const uint BulletPictureMask = 1U << 23;
        private const uint BulletSchemeMask = 1U << 24;
        private const uint BulletHasSchemeMask = 1U << 25;
        private const uint AllowedParagraphMasks = BulletPictureMask
            | BulletSchemeMask | BulletHasSchemeMask;
        private const uint CharacterPpt10ExtensionMask = 1U << 20;

        internal static LegacyPptTextBody Apply(LegacyPptTextBody textBody,
            LegacyPptRecord? styleRecord, bool malformedContainer = false) {
            if (textBody == null) throw new ArgumentNullException(
                nameof(textBody));
            if (styleRecord == null) {
                return malformedContainer
                    ? textBody.WithPpt9Formatting(textBody.ParagraphRuns,
                        hasUnprojectedParagraphFormatting: true,
                        isStyle9Truncated: true)
                    : textBody;
            }
            try {
                if (styleRecord.Version != 0 || styleRecord.Instance != 0
                    || styleRecord.Type != 0x0FAC) {
                    throw new InvalidDataException(
                        "The PPT9 extended-style atom has an invalid record header.");
                }
                var cursor = new LegacyPptTextPropertyCursor(styleRecord,
                    "StyleTextProp9Atom");
                var entries = new List<Style9Entry>();
                while (!cursor.IsAtEnd) entries.Add(ReadEntry(cursor));
                return ApplyEntries(textBody, entries,
                    malformedContainer);
            } catch (Exception exception) when (exception
                is InvalidDataException or OverflowException
                    or ArgumentOutOfRangeException) {
                return textBody.WithPpt9Formatting(textBody.ParagraphRuns,
                    hasUnprojectedParagraphFormatting: true,
                    isStyle9Truncated: true);
            }
        }

        private static Style9Entry ReadEntry(
            LegacyPptTextPropertyCursor cursor) {
            uint masks = cursor.ReadUInt32();
            if ((masks & ~AllowedParagraphMasks) != 0) {
                throw new InvalidDataException(
                    "A TextPFException9 uses fields outside the PPT9 extended paragraph contract.");
            }
            ushort? pictureReference = (masks & BulletPictureMask) != 0
                ? cursor.ReadUInt16()
                : null;
            bool? hasAutoNumber = null;
            if ((masks & BulletHasSchemeMask) != 0) {
                short value = cursor.ReadInt16();
                if (value is not 0 and not 1) {
                    throw new InvalidDataException(
                        "A TextPFException9 automatic-numbering flag is invalid.");
                }
                hasAutoNumber = value == 1;
            }
            LegacyPptAutoNumberScheme? scheme = null;
            short? startAt = null;
            if ((masks & BulletSchemeMask) != 0) {
                ushort value = cursor.ReadUInt16();
                if (value > (ushort)LegacyPptAutoNumberScheme
                        .HindiAlpha1Period) {
                    throw new InvalidDataException(
                        "A TextPFException9 automatic-numbering scheme is undefined.");
                }
                scheme = (LegacyPptAutoNumberScheme)value;
                startAt = cursor.ReadInt16();
                if (startAt < 1) {
                    throw new InvalidDataException(
                        "A TextPFException9 automatic-numbering start is less than one.");
                }
            }
            if (hasAutoNumber == true && !scheme.HasValue) {
                scheme = LegacyPptAutoNumberScheme.ArabicPeriod;
                startAt = 1;
            }

            uint characterMasks = cursor.ReadUInt32();
            if ((characterMasks & ~CharacterPpt10ExtensionMask) != 0) {
                throw new InvalidDataException(
                    "A TextCFException9 contains fields forbidden by StyleTextProp9.");
            }
            bool hasUnprojectedFormatting = pictureReference.HasValue
                || characterMasks != 0;
            if ((characterMasks & CharacterPpt10ExtensionMask) != 0) {
                cursor.ReadUInt32();
            }
            uint specialInfoMasks = cursor.ReadUInt32();
            if (specialInfoMasks != 0) {
                throw new InvalidDataException(
                    "A StyleTextProp9 TextSIException contains forbidden fields.");
            }
            if (hasAutoNumber == false && (scheme.HasValue
                    || startAt.HasValue)) {
                hasUnprojectedFormatting = true;
            }
            return new Style9Entry(hasAutoNumber, scheme, startAt,
                pictureReference, hasUnprojectedFormatting);
        }

        private static LegacyPptTextBody ApplyEntries(
            LegacyPptTextBody textBody, IReadOnlyList<Style9Entry> entries,
            bool malformedContainer) {
            if (entries.Count == 0) {
                return textBody.WithPpt9Formatting(textBody.ParagraphRuns,
                    hasUnprojectedParagraphFormatting:
                        malformedContainer,
                    isStyle9Truncated: malformedContainer);
            }
            var formattingByParagraph = new Dictionary<int, Style9Entry>();
            IReadOnlyList<CharacterGroup> groups = CreateCharacterGroups(
                textBody);
            int entryIndex = 0;
            bool linkageFailed = false;
            foreach (CharacterGroup group in groups) {
                while (entryIndex < entries.Count
                       && entryIndex % 16 != group.Ppt9RunId) {
                    entryIndex++;
                }
                if (entryIndex >= entries.Count) {
                    linkageFailed = true;
                    break;
                }
                Style9Entry entry = entries[entryIndex++];
                foreach (LegacyPptParagraphRun paragraph
                         in textBody.ParagraphRuns) {
                    if (paragraph.Start >= group.Start
                        && paragraph.Start < group.End) {
                        formattingByParagraph[paragraph.Start] = entry;
                    }
                }
            }

            var paragraphs = new List<LegacyPptParagraphRun>(
                textBody.ParagraphRuns.Count);
            bool hasUnprojectedFormatting = malformedContainer
                || linkageFailed;
            foreach (LegacyPptParagraphRun paragraph
                     in textBody.ParagraphRuns) {
                if (!formattingByParagraph.TryGetValue(paragraph.Start,
                        out Style9Entry? entry)) {
                    paragraphs.Add(paragraph);
                    continue;
                }
                hasUnprojectedFormatting |= entry
                    .HasUnprojectedFormatting;
                paragraphs.Add(paragraph.WithPpt9Formatting(
                    entry.HasAutoNumber, entry.AutoNumberScheme,
                    entry.AutoNumberStartAt, entry.BulletPictureReference,
                    entry.HasUnprojectedFormatting));
            }
            return textBody.WithPpt9Formatting(paragraphs,
                hasUnprojectedFormatting,
                isStyle9Truncated: malformedContainer || linkageFailed);
        }

        private static IReadOnlyList<CharacterGroup> CreateCharacterGroups(
            LegacyPptTextBody textBody) {
            var groups = new List<CharacterGroup>();
            foreach (LegacyPptCharacterRun run in textBody.CharacterRuns) {
                byte runId = run.Ppt9RunId ?? 0;
                int end = checked(run.Start + run.Length);
                if (groups.Count != 0
                    && groups[groups.Count - 1].Ppt9RunId == runId
                    && groups[groups.Count - 1].End == run.Start) {
                    groups[groups.Count - 1] = new CharacterGroup(
                        groups[groups.Count - 1].Start, end, runId);
                } else {
                    groups.Add(new CharacterGroup(run.Start, end, runId));
                }
            }
            if (groups.Count == 0 && textBody.ParagraphRuns.Count != 0) {
                int start = textBody.ParagraphRuns.Min(run => run.Start);
                int end = textBody.ParagraphRuns.Max(run => checked(
                    run.Start + Math.Max(run.Length, 1)));
                groups.Add(new CharacterGroup(start, end, 0));
            }
            return groups;
        }

        private sealed class Style9Entry {
            internal Style9Entry(bool? hasAutoNumber,
                LegacyPptAutoNumberScheme? autoNumberScheme,
                short? autoNumberStartAt, ushort? bulletPictureReference,
                bool hasUnprojectedFormatting) {
                HasAutoNumber = hasAutoNumber;
                AutoNumberScheme = autoNumberScheme;
                AutoNumberStartAt = autoNumberStartAt;
                BulletPictureReference = bulletPictureReference;
                HasUnprojectedFormatting = hasUnprojectedFormatting;
            }

            internal bool? HasAutoNumber { get; }

            internal LegacyPptAutoNumberScheme? AutoNumberScheme { get; }

            internal short? AutoNumberStartAt { get; }

            internal ushort? BulletPictureReference { get; }

            internal bool HasUnprojectedFormatting { get; }
        }

        private readonly struct CharacterGroup {
            internal CharacterGroup(int start, int end, byte ppt9RunId) {
                Start = start;
                End = end;
                Ppt9RunId = ppt9RunId;
            }

            internal int Start { get; }

            internal int End { get; }

            internal byte Ppt9RunId { get; }
        }
    }
}
