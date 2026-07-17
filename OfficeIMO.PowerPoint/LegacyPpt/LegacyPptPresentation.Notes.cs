using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private IReadOnlyDictionary<uint, LegacyPptNotesDirectoryEntry> ReadNotesDirectory(
            LegacyPptRecord document, LegacyPptImportOptions options) {
            LegacyPptRecord? list = document.Children.FirstOrDefault(record =>
                record.Type == RecordSlideListWithText && record.Instance == 2);
            if (list == null) return new Dictionary<uint, LegacyPptNotesDirectoryEntry>();

            var result = new Dictionary<uint, LegacyPptNotesDirectoryEntry>();
            foreach (LegacyPptRecord atom in list.Children.Where(record =>
                         record.Type == RecordSlidePersistAtom)) {
                if (atom.PayloadLength != 20) {
                    AddDiagnostic("PPT-NOTES-PERSIST-LENGTH",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"A NotesPersistAtom has {atom.PayloadLength} payload bytes instead of 20 and remains preserve-only.",
                        atom.Offset);
                    continue;
                }
                uint persistId = atom.ReadUInt32(0);
                uint flags = atom.ReadUInt32(4);
                uint notesId = atom.ReadUInt32(12);
                if (persistId == 0 || notesId == 0 || result.ContainsKey(notesId)) {
                    AddDiagnostic("PPT-NOTES-PERSIST-REFERENCE",
                        LegacyPptDiagnosticSeverity.Warning,
                        "A NotesPersistAtom has a null or duplicate notes/persist identifier and remains preserve-only.",
                        atom.Offset);
                    continue;
                }
                if (options.ReportUnsupportedContent
                    && ((flags & ~0x00000004U) != 0 || atom.ReadUInt32(8) != 0
                        || atom.ReadUInt32(16) != 0)) {
                    AddDiagnostic("PPT-NOTES-PERSIST-RESERVED",
                        LegacyPptDiagnosticSeverity.Warning,
                        "A NotesPersistAtom uses reserved fields; those values remain preserved only.",
                        atom.Offset);
                }
                result.Add(notesId, new LegacyPptNotesDirectoryEntry(notesId, persistId,
                    (flags & 0x00000004U) != 0));
            }
            return result;
        }

        private void TryReadNotes(LegacyPptSlide slide, byte[] documentStream,
            IReadOnlyDictionary<uint, uint> persistOffsets,
            IReadOnlyDictionary<uint, LegacyPptNotesDirectoryEntry> notesDirectory,
            LegacyPptImportOptions options) {
            if (slide.NotesId == 0) return;
            if (!notesDirectory.TryGetValue(slide.NotesId,
                    out LegacyPptNotesDirectoryEntry directoryEntry)) {
                AddDiagnostic("PPT-NOTES-ID-MISSING", LegacyPptDiagnosticSeverity.Warning,
                    $"Slide {slide.SlideId} references missing notes id {slide.NotesId}.",
                    offset: null);
                return;
            }
            if (!persistOffsets.TryGetValue(directoryEntry.PersistId, out uint notesOffset)) {
                AddDiagnostic("PPT-NOTES-PERSIST-MISSING", LegacyPptDiagnosticSeverity.Warning,
                    $"Notes id {slide.NotesId} references missing persist object {directoryEntry.PersistId}.",
                    offset: null);
                return;
            }

            try {
                LegacyPptRecord notes = LegacyPptRecordReader.ReadSingle(documentStream,
                    ToBoundedOffset(notesOffset, documentStream.Length,
                        "notes persist object"), options, _recordBudget);
                if (notes.Type != RecordNotes) {
                    AddDiagnostic("PPT-NOTES-TYPE", LegacyPptDiagnosticSeverity.Warning,
                        $"Notes id {slide.NotesId} points to record 0x{notes.Type:X4} instead of a NotesContainer.",
                        notes.Offset);
                    return;
                }

                LegacyPptRecord? atom = notes.Children.FirstOrDefault(record =>
                    record.Type == RecordNotesAtom);
                if (atom == null || atom.PayloadLength != 8) {
                    AddDiagnostic("PPT-NOTES-ATOM-LENGTH", LegacyPptDiagnosticSeverity.Warning,
                        "A notes page has no complete 8-byte NotesAtom and remains preserve-only.",
                        atom?.Offset ?? notes.Offset);
                    return;
                }
                uint referencedSlideId = atom.ReadUInt32(0);
                ushort flags = atom.ReadUInt16(4);
                var page = new LegacyPptNotesPage(slide.NotesId,
                    directoryEntry.PersistId, referencedSlideId) {
                    FollowsMasterObjects = (flags & 0x0001) != 0,
                    FollowsMasterColorScheme = (flags & 0x0002) != 0,
                    FollowsMasterBackground = (flags & 0x0004) != 0,
                    ColorScheme = ReadColorScheme(notes),
                    RoundTripTheme = ReadRoundTripTheme(notes,
                        $"notes page {slide.NotesId}", options)
                };
                if (referencedSlideId != slide.SlideId) {
                    AddDiagnostic("PPT-NOTES-SLIDE-MISMATCH",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"Notes id {slide.NotesId} references slide {referencedSlideId} instead of {slide.SlideId}.",
                        atom.Offset);
                }
                if (options.ReportUnsupportedContent
                    && ((flags & 0xFFF8) != 0 || atom.ReadUInt16(6) != 0)) {
                    AddDiagnostic("PPT-NOTES-ATOM-RESERVED",
                        LegacyPptDiagnosticSeverity.Warning,
                        "A NotesAtom uses reserved bits; those values remain preserved only.",
                        atom.Offset);
                }

                LegacyPptColorScheme? effectiveScheme = page.FollowsMasterColorScheme
                    ? NotesMaster?.ColorScheme
                    : page.ColorScheme;
                page.Background = ReadBackground(notes,
                    effectiveScheme ?? page.ColorScheme, options);
                ParseShapes(notes, page.AddShape, "notes page", options,
                    effectiveScheme ?? page.ColorScheme, page.AddConnectorRule);
                page.Text = ReadNotesBodyText(page.Shapes, notes);
                slide.NotesPage = page;
                slide.NotesText = page.Text;
            } catch (InvalidDataException exception) {
                AddDiagnostic("PPT-NOTES-READ", LegacyPptDiagnosticSeverity.Warning,
                    $"Speaker notes could not be decoded: {exception.Message}", notesOffset);
            }
        }

        private string ReadNotesBodyText(IReadOnlyList<LegacyPptShape> shapes,
            LegacyPptRecord notesRecord) {
            LegacyPptShape[] flattened = FlattenNotesShapes(shapes).ToArray();
            LegacyPptShape[] bodyShapes = flattened.Where(shape =>
                    shape.PlaceholderKind is LegacyPptPlaceholderKind.NotesBody
                        or LegacyPptPlaceholderKind.MasterNotesBody)
                .ToArray();
            if (bodyShapes.Length == 0) {
                bodyShapes = flattened.Where(shape =>
                        shape.TextBody.TextType == LegacyPptTextType.Notes)
                    .ToArray();
            }
            string text = string.Join("\n", bodyShapes
                .Select(shape => shape.Text)
                .Where(value => !string.IsNullOrWhiteSpace(value)));
            if (!string.IsNullOrWhiteSpace(text)) return text;

            return string.Join("\n", notesRecord.DescendantsAndSelf()
                .Where(record => record.Type == OfficeArtClientTextbox)
                .Select(record => ReadText(record).Text)
                .Where(value => !string.IsNullOrWhiteSpace(value)));
        }

        private static IEnumerable<LegacyPptShape> FlattenNotesShapes(
            IEnumerable<LegacyPptShape> shapes) {
            foreach (LegacyPptShape shape in shapes) {
                yield return shape;
                foreach (LegacyPptShape child in FlattenNotesShapes(shape.Children)) {
                    yield return child;
                }
            }
        }

        private readonly struct LegacyPptNotesDirectoryEntry {
            internal LegacyPptNotesDirectoryEntry(uint notesId, uint persistId,
                bool hasNonOutlineData) {
                NotesId = notesId;
                PersistId = persistId;
                HasNonOutlineData = hasNonOutlineData;
            }

            internal uint NotesId { get; }
            internal uint PersistId { get; }
            internal bool HasNonOutlineData { get; }
        }
    }
}
