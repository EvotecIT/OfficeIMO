using System.Threading;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Admits old in-stream note PLCs while retaining global source character positions.</summary>
    internal static class LegacyWord69Notes {
        /// <summary>Maps the old common DOP note settings through the shared section model.</summary>
        internal static IReadOnlyList<LegacyDocSection> ApplySettings(
            LegacyWord69Bytes dop, IReadOnlyList<LegacyDocSection> sections, bool hasFootnotes, bool hasEndnotes,
            Action<LegacyDocUnsupportedFeatureKind, string, string> report) {
            if (!hasFootnotes && !hasEndnotes) return sections;
            if (dop.Length < 0x38) {
                report(LegacyDocUnsupportedFeatureKind.Section, "doc-word69-note-settings",
                    "The Word 6/95 note stories have no complete common DOP note settings; their numbering and placement are unqualified.");
                return sections;
            }
            ushort footnoteFlags = LegacyWord69Reader.U16(dop, 2);
            ushort endnoteFlags = LegacyWord69Reader.U16(dop, 0x34);
            ushort formats = LegacyWord69Reader.U16(dop, 0x36);
            FootnotePositionValues? footnotePosition = ((dop[0] >> 5) & 3) switch {
                1 => FootnotePositionValues.PageBottom,
                2 => FootnotePositionValues.BeneathText,
                _ => null
            };
            EndnotePositionValues? endnotePosition = (formats & 3) switch {
                0 => EndnotePositionValues.SectionEnd,
                3 => EndnotePositionValues.DocumentEnd,
                _ => null
            };
            RestartNumberValues? footnoteRestart = ReadRestart(footnoteFlags & 3);
            RestartNumberValues? endnoteRestart = ReadRestart(endnoteFlags & 3);
            if (endnoteRestart == RestartNumberValues.EachPage) endnoteRestart = null;
            int footnoteStart = footnoteFlags >> 2;
            int endnoteStart = endnoteFlags >> 2;
            NumberFormatValues? footnoteFormat = LegacyDocNumberFormatMapper.FromNfc((byte)((formats >> 2) & 15));
            NumberFormatValues? endnoteFormat = LegacyDocNumberFormatMapper.FromNfc((byte)((formats >> 6) & 15));
            if (hasFootnotes && (footnotePosition == null || footnoteRestart == null || footnoteStart == 0 || footnoteFormat == null)
                || hasEndnotes && (endnotePosition == null || endnoteRestart == null || endnoteStart == 0 || endnoteFormat == null)) {
                report(LegacyDocUnsupportedFeatureKind.Section, "doc-word69-note-settings",
                    "The Word 6/95 DOP contains note numbering or placement outside the qualified editable note settings.");
            }
            return sections.Select(section => {
                LegacyDocSectionFormat format = section.Format;
                return new LegacyDocSection(section.StartCharacter, section.EndCharacter, format.WithNoteSettings(
                    footnotePosition: hasFootnotes ? format.FootnotePosition ?? footnotePosition : null,
                    footnoteRestart: hasFootnotes ? format.FootnoteRestart ?? footnoteRestart : null,
                    footnoteStart: hasFootnotes ? format.FootnoteStart ?? (footnoteStart == 0 ? null : footnoteStart) : null,
                    footnoteNumberFormat: hasFootnotes ? format.FootnoteNumberFormat ?? footnoteFormat : null,
                    endnotePosition: hasEndnotes ? format.EndnotePosition ?? endnotePosition : null,
                    endnoteRestart: hasEndnotes ? format.EndnoteRestart ?? endnoteRestart : null,
                    endnoteStart: hasEndnotes ? format.EndnoteStart ?? (endnoteStart == 0 ? null : endnoteStart) : null,
                    endnoteNumberFormat: hasEndnotes ? format.EndnoteNumberFormat ?? endnoteFormat : null));
            }).ToArray();

            static RestartNumberValues? ReadRestart(int value) => value switch {
                0 => RestartNumberValues.Continuous,
                1 => RestartNumberValues.EachSection,
                2 => RestartNumberValues.EachPage,
                _ => null
            };
        }

        internal static void Read(
            LegacyWord69Bytes bytes, int usedEnd, int textStart, LegacyDocTextCharacter[] characters,
            int bodyLength, int footnoteLength, int endnoteStart, int endnoteLength,
            IReadOnlyList<LegacyDocCharacterFormatRange> characterRanges,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphRanges,
            LegacyDocFieldMetadataReader fieldMetadata,
            Action<LegacyDocUnsupportedFeatureKind, string, string> report,
            out IReadOnlyList<LegacyDocFootnote> footnotes,
            out IReadOnlyList<LegacyDocEndnote> endnotes,
            CancellationToken cancellationToken = default) {
            var anchors = new HashSet<int>();
            var footnoteData = ReadKind(footnoteLength, bodyLength, 0x68, 0x70, false);
            var endnoteData = ReadKind(endnoteLength, endnoteStart, 0x1D2, 0x1DA, true);
            var projectedFootnotes = new LegacyDocFootnote[footnoteData.Count];
            for (int index = 0; index < projectedFootnotes.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                var note = footnoteData[index];
                projectedFootnotes[index] = new LegacyDocFootnote(note.Reference, note.Paragraphs);
            }
            var projectedEndnotes = new LegacyDocEndnote[endnoteData.Count];
            for (int index = 0; index < projectedEndnotes.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                var note = endnoteData[index];
                projectedEndnotes[index] = new LegacyDocEndnote(note.Reference, note.Paragraphs);
            }
            footnotes = projectedFootnotes;
            endnotes = projectedEndnotes;
            for (int index = 0; index < bodyLength; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (characters[index].Character == LegacyDocFootnoteReader.FootnoteReferenceCharacter
                    && !anchors.Contains(index)) {
                    report(LegacyDocUnsupportedFeatureKind.Formatting, "doc-word69-note-anchor-unresolved",
                        "A Word 6/95 body note marker has no matching qualified note reference.");
                    break;
                }
            }

            List<(int Reference, IReadOnlyList<LegacyDocNoteParagraph> Paragraphs)> ReadKind(
                int storyLength, int storyStart, int referenceDirectory, int textDirectory, bool endnote) {
                var notes = new List<(int, IReadOnlyList<LegacyDocNoteParagraph>)>();
                if (storyLength == 0) return notes;
                cancellationToken.ThrowIfCancellationRequested();
                string name = endnote ? "endnote" : "footnote";
                var kind = endnote ? LegacyDocUnsupportedFeatureKind.Endnote : LegacyDocUnsupportedFeatureKind.Footnote;
                string code = endnote ? "doc-word69-endnotes" : "doc-word69-footnotes";
                // The old minimum FIB precedes the optional endnote/field directories.
                int requiredHeader = endnote ? 0x1EA : LegacyWord69Reader.FibSize;
                if (textStart < requiredHeader) {
                    throw new InvalidDataException($"The Word 6/95 {name} directories overlap its text.");
                }
                LegacyWord69Bytes references = ReadDirectory(referenceDirectory, requiredHeader);
                LegacyWord69Bytes positions = ReadDirectory(textDirectory, requiredHeader);
                if (references.IsEmpty || positions.IsEmpty) {
                    report(kind, code, $"The Word 6/95 {name} story has no matching reference and text PLCs.");
                    return notes;
                }
                if (references.Length < 10 || (references.Length - 4) % 6 != 0
                    || positions.Length < 8 || positions.Length % 4 != 0) {
                    throw new InvalidDataException($"The Word 6/95 {name} PLCs have an invalid record length.");
                }
                int count = (references.Length - 4) / 6;
                int positionCount = positions.Length / 4;
                if (positionCount < count + 1) {
                    throw new InvalidDataException($"The Word 6/95 {name} PLCs omit a declared text range.");
                }
                int previousText = -1;
                for (int index = 0; index < positionCount; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int position = LegacyWord69Reader.I32(positions, index * 4);
                    // Only the reference-indexed n+1 CPs delimit note text. Old Word
                    // can retain one further PLC position beyond the decoded story.
                    if (position < previousText || position < 0 || (index <= count && position > storyLength)) {
                        throw new InvalidDataException($"The Word 6/95 {name} text PLC exceeds its story.");
                    }
                    previousText = position;
                }
                int descriptorStart = (count + 1) * 4;
                int previousReference = -1;
                int unprojected = 0;
                bool unqualifiedContent = false;
                for (int index = 0; index < count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int reference = LegacyWord69Reader.I32(references, index * 4);
                    if (reference <= previousReference || reference < 0 || reference >= bodyLength
                        || !anchors.Add(reference)) {
                        throw new InvalidDataException($"The Word 6/95 {name} references are duplicate, unordered or outside the body.");
                    }
                    previousReference = reference;
                    int start = LegacyWord69Reader.I32(positions, index * 4);
                    int end = LegacyWord69Reader.I32(positions, (index + 1) * 4);
                    bool automatic = LegacyWord69Reader.U16(references, descriptorStart + index * 2) != 0;
                    LegacyDocTextCharacter anchor = characters[reference];
                    if (!automatic || (char.IsControl(anchor.Character) && anchor.Character != LegacyDocFootnoteReader.FootnoteReferenceCharacter)
                        || end == start) {
                        unprojected++;
                        continue;
                    }
                    var noteCharacters = new ArraySegment<LegacyDocTextCharacter>(characters, storyStart + start, end - start);
                    var noteParagraphs = LegacyDocFootnoteReader.BuildStoryParagraphs(
                        noteCharacters, storyStart + start, storyStart + end, characterRanges, paragraphRanges,
                        new LegacyDocBookmarkProjectionTracker(Array.Empty<LegacyDocBookmark>()),
                        new Dictionary<int, LegacyDocPicture>(), fieldMetadata,
                        orderedFormattingRanges: true, preserveExplicitEmptyParagraphs: true, cancellationToken: cancellationToken);
                    int contentStart = start + (noteCharacters[0].Character == LegacyDocFootnoteReader.FootnoteReferenceCharacter ? 1 : 0);
                    for (int position = contentStart; position < end; position++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        // A flat note paragraph model cannot retain nested anchors or table cells.
                        if (characters[storyStart + position].Character is '\u0002' or '\a') unqualifiedContent = true;
                    }
                    foreach (LegacyDocNoteParagraph paragraph in noteParagraphs) {
                        cancellationToken.ThrowIfCancellationRequested();
                        unqualifiedContent |= paragraph.Format.IsInTable == true;
                    }
                    if (noteParagraphs.Count == 0) {
                        unprojected++;
                    } else {
                        // Word95 can store a printable cached number at an automatic
                        // reference. Normalize its meaning without moving source CP/FC.
                        if (anchor.Character != LegacyDocFootnoteReader.FootnoteReferenceCharacter) {
                            characters[reference] = new LegacyDocTextCharacter(
                                LegacyDocFootnoteReader.FootnoteReferenceCharacter, anchor.FileOffset, anchor.CharacterPosition);
                        }
                        notes.Add((reference, noteParagraphs));
                    }
                }
                int terminalReference = LegacyWord69Reader.I32(references, count * 4);
                if (terminalReference <= previousReference) {
                    throw new InvalidDataException($"The Word 6/95 {name} reference PLC has an unordered terminal position.");
                }
                if (unprojected != 0) {
                    report(kind, code,
                        $"The Word 6/95 {name} story has {unprojected} custom, unmatched or empty note range(s) outside the qualified automatic-note projection.");
                }
                if (unqualifiedContent) {
                    report(kind, "doc-word69-note-content-unqualified",
                        $"The Word 6/95 {name} story contains nested note markers or table structure outside the editable paragraph-note projection.");
                }
                int firstText = LegacyWord69Reader.I32(positions, 0);
                int consumedEnd = LegacyWord69Reader.I32(positions, count * 4);
                bool unconsumed = HasContent(0, firstText) || HasContent(consumedEnd, storyLength);
                if (positionCount > count + 2 || unconsumed) {
                    report(kind, code, $"The Word 6/95 {name} text PLC leaves unqualified story content or extra ranges.");
                }
                return notes;

                bool HasContent(int start, int end) {
                    for (int position = start; position < end; position++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (characters[storyStart + position].Character != '\r') return true;
                    }
                    return false;
                }
            }

            LegacyWord69Bytes ReadDirectory(int directory, int minimumHeader) {
                int length = LegacyWord69Reader.I32(bytes, directory + 4);
                if (length != 0 && LegacyWord69Reader.I32(bytes, directory) < minimumHeader) {
                    throw new InvalidDataException("A Word 6/95 note PLC overlaps its FIB.");
                }
                return LegacyWord69Reader.ReadRange(bytes, usedEnd, directory);
            }
        }
    }
}
