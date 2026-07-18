using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal static bool ShouldWriteNotesPage(PowerPointSlide slide,
            out string text) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            text = slide.Notes.TryGetText(out string noteText)
                ? noteText
                : string.Empty;
            if (!string.IsNullOrWhiteSpace(text)) return true;

            NotesSlidePart? notesPart = slide.SlidePart.NotesSlidePart;
            return notesPart?.NotesSlide?.CommonSlideData?.Background != null
                || notesPart?.ThemeOverridePart?.ThemeOverride != null
                || notesPart?.NotesSlide?.ShowMasterShapes?.Value == false;
        }

        private static byte[] BuildNotesList(IReadOnlyList<LegacyPptWriterNote> notes) {
            var children = new List<byte[]>(notes.Count);
            foreach (LegacyPptWriterNote note in notes) {
                var payload = new byte[20];
                WriteUInt32(payload, 0, note.PersistId);
                WriteUInt32(payload, 4, 4U);
                WriteUInt32(payload, 12, note.NotesId);
                children.Add(BuildRecord(version: 0, instance: 0,
                    RecordSlidePersistAtom, payload));
            }
            return BuildContainer(RecordSlideListWithText, instance: 2, children);
        }

        private static byte[] BuildNotesRecord(LegacyPptRecord prototype, string text,
            uint slideId, uint drawingId, NotesSlidePart? sourcePart,
            LegacyPptWriterPictureCatalog pictureCatalog) {
            var children = new List<byte[]>(prototype.Children.Count);
            A.ThemeOverride? theme = sourcePart?.ThemeOverridePart?
                .ThemeOverride;
            IReadOnlyList<byte[]> roundTripThemeRecords =
                BuildRoundTripThemeRecords(theme,
                    sourcePart?.NotesSlide?.ColorMapOverride);
            A.ColorScheme? overrideColors = theme?.ColorScheme;
            LegacyPptWriterColorScheme? classicOverride = overrideColors == null
                ? null
                : ReadColorScheme(overrideColors);
            LegacyPptWriterBackground? background = null;
            if (sourcePart != null
                && !TryReadBackground(sourcePart, out background,
                    out string? backgroundReason)) {
                throw new NotSupportedException(backgroundReason);
            }
            bool replacedNotesBody = false;
            bool wroteClassicOverride = false;
            foreach (LegacyPptRecord child in prototype.Children) {
                if (child.Type == RecordNotesAtom) {
                    byte[] atom = child.CopyRecordBytes();
                    WriteUInt32(atom, 8, slideId);
                    ushort flags = ReadUInt16(atom, 12);
                    flags = background == null
                        ? unchecked((ushort)(flags | 0x0004))
                        : unchecked((ushort)(flags & ~0x0004));
                    flags = classicOverride == null
                        ? unchecked((ushort)(flags | 0x0002))
                        : unchecked((ushort)(flags & ~0x0002));
                    flags = sourcePart?.NotesSlide?.ShowMasterShapes?.Value != false
                        ? unchecked((ushort)(flags | 0x0001))
                        : unchecked((ushort)(flags & ~0x0001));
                    WriteUInt16(atom, 12, flags);
                    children.Add(atom);
                } else if (child.Type == RecordDrawing) {
                    byte[] drawingBytes = RewriteNotesDrawingRecord(child,
                        text, drawingId, replaceNotesBody: false,
                        ref replacedNotesBody);
                    if (background != null) {
                        LegacyPptRecord drawingRecord = LegacyPptRecordReader
                            .ReadSingle(drawingBytes, 0,
                                new LegacyPptImportOptions());
                        drawingBytes = BuildBackgroundDrawingRecord(
                            drawingRecord, background, pictureCatalog);
                    }
                    children.Add(drawingBytes);
                } else if (classicOverride != null
                           && child.Type == RecordColorSchemeAtom
                           && child.Instance == 1) {
                    children.Add(BuildColorSchemeAtom(classicOverride));
                    wroteClassicOverride = true;
                } else if (!IsRoundTripThemeRecord(child.Type)) {
                    children.Add(child.CopyRecordBytes());
                }
            }
            if (!replacedNotesBody) {
                throw new InvalidDataException(
                    "The embedded PowerPoint notes template has no notes-body text box.");
            }
            if (classicOverride != null && !wroteClassicOverride) {
                children.Add(BuildColorSchemeAtom(classicOverride));
            }
            children.AddRange(roundTripThemeRecords);
            return BuildContainer(RecordNotes, instance: 0, children);
        }

        private static byte[] RewriteNotesDrawingRecord(LegacyPptRecord record, string text,
            uint drawingId, bool replaceNotesBody, ref bool replacedNotesBody) {
            if (record.Type == OfficeArtSpContainer) {
                LegacyPptRecord? placeholder = record.DescendantsAndSelf().FirstOrDefault(child =>
                    child.Type == RecordPlaceholder && child.PayloadLength >= 5);
                replaceNotesBody = placeholder?.ReadByte(4) is 0x06 or 0x0C;
            }
            if (record.Type == OfficeArtClientTextbox && replaceNotesBody) {
                replacedNotesBody = true;
                return BuildTextBox(text, textType: 2U);
            }
            if (record.Version == 0x0F && record.Children.Count > 0) {
                var children = new List<byte[]>(record.Children.Count);
                foreach (LegacyPptRecord child in record.Children) {
                    children.Add(RewriteNotesDrawingRecord(child, text, drawingId,
                        replaceNotesBody, ref replacedNotesBody));
                }
                return BuildRecord(record.Version, record.Instance, record.Type,
                    Concat(children));
            }

            byte[] bytes = record.CopyRecordBytes();
            if (record.Type == OfficeArtDg) {
                WriteUInt16(bytes, 0, checked((ushort)((drawingId << 4) | record.Version)));
                uint sourceCurrentShapeId = ReadUInt32(bytes, 12);
                WriteUInt32(bytes, 12, checked((drawingId << 10)
                    | (sourceCurrentShapeId & 0x000003FFU)));
            } else if (record.Type == OfficeArtFsp) {
                uint sourceShapeId = ReadUInt32(bytes, 8);
                WriteUInt32(bytes, 8, checked((drawingId << 10)
                    | (sourceShapeId & 0x000003FFU)));
            }
            return bytes;
        }

        private sealed class LegacyPptWriterNote {
            internal LegacyPptWriterNote(int slideIndex, string text, uint notesId,
                uint persistId, uint drawingId, NotesSlidePart? sourcePart) {
                SlideIndex = slideIndex;
                Text = text ?? string.Empty;
                NotesId = notesId;
                PersistId = persistId;
                DrawingId = drawingId;
                SourcePart = sourcePart;
            }

            internal int SlideIndex { get; }
            internal string Text { get; }
            internal uint NotesId { get; }
            internal uint PersistId { get; }
            internal uint DrawingId { get; }
            internal NotesSlidePart? SourcePart { get; }
        }
    }
}
