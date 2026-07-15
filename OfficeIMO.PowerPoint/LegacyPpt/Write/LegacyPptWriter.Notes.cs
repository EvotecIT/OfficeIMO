using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
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
            uint slideId, uint drawingId) {
            var children = new List<byte[]>(prototype.Children.Count);
            bool replacedNotesBody = false;
            foreach (LegacyPptRecord child in prototype.Children) {
                if (child.Type == RecordNotesAtom) {
                    byte[] atom = child.CopyRecordBytes();
                    WriteUInt32(atom, 8, slideId);
                    children.Add(atom);
                } else if (child.Type == RecordDrawing) {
                    children.Add(RewriteNotesDrawingRecord(child, text, drawingId,
                        replaceNotesBody: false, ref replacedNotesBody));
                } else {
                    children.Add(child.CopyRecordBytes());
                }
            }
            if (!replacedNotesBody) {
                throw new InvalidDataException(
                    "The embedded PowerPoint notes template has no notes-body text box.");
            }
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
                uint persistId, uint drawingId) {
                SlideIndex = slideIndex;
                Text = text ?? string.Empty;
                NotesId = notesId;
                PersistId = persistId;
                DrawingId = drawingId;
            }

            internal int SlideIndex { get; }
            internal string Text { get; }
            internal uint NotesId { get; }
            internal uint PersistId { get; }
            internal uint DrawingId { get; }
        }
    }
}
