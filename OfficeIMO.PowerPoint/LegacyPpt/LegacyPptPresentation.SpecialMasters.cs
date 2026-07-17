using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordNotes = 0x03F0;
        private const ushort RecordNotesAtom = 0x03F1;
        private const ushort RecordHandout = 0x0FC9;

        private LegacyPptSpecialMaster? _notesMaster;
        private LegacyPptSpecialMaster? _handoutMaster;

        /// <summary>Gets the notes master referenced by the document atom, when present.</summary>
        public LegacyPptSpecialMaster? NotesMaster => _notesMaster;

        /// <summary>Gets the handout master referenced by the document atom, when present.</summary>
        public LegacyPptSpecialMaster? HandoutMaster => _handoutMaster;

        private void ParseSpecialMasters(LegacyPptRecord? documentAtom, byte[] documentStream,
            IReadOnlyDictionary<uint, uint> persistOffsets, LegacyPptImportOptions options) {
            if (documentAtom == null) return;
            if (DocumentSettings == null) return;
            _notesMaster = ReadSpecialMaster(DocumentSettings.NotesMasterPersistId,
                LegacyPptSpecialMasterKind.Notes, RecordNotes, documentStream, persistOffsets, options);
            _handoutMaster = ReadSpecialMaster(DocumentSettings.HandoutMasterPersistId,
                LegacyPptSpecialMasterKind.Handout, RecordHandout, documentStream, persistOffsets, options);
        }

        private LegacyPptSpecialMaster? ReadSpecialMaster(uint persistId,
            LegacyPptSpecialMasterKind kind, ushort expectedRecordType, byte[] documentStream,
            IReadOnlyDictionary<uint, uint> persistOffsets, LegacyPptImportOptions options) {
            if (persistId == 0) return null;
            if (!persistOffsets.TryGetValue(persistId, out uint offset)) {
                AddDiagnostic("PPT-SPECIAL-MASTER-PERSIST-MISSING",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"The {GetSpecialMasterName(kind)} master references missing persist object {persistId}.",
                    offset: null);
                return null;
            }

            LegacyPptRecord record;
            try {
                record = LegacyPptRecordReader.ReadSingle(documentStream,
                    ToBoundedOffset(offset, documentStream.Length,
                        $"{GetSpecialMasterName(kind)} master persist object"),
                    options, _recordBudget);
            } catch (InvalidDataException exception) {
                AddDiagnostic("PPT-SPECIAL-MASTER-READ", LegacyPptDiagnosticSeverity.Warning,
                    $"The {GetSpecialMasterName(kind)} master could not be decoded: {exception.Message}",
                    offset);
                return null;
            }
            if (record.Type != expectedRecordType) {
                AddDiagnostic("PPT-SPECIAL-MASTER-TYPE", LegacyPptDiagnosticSeverity.Warning,
                    $"The {GetSpecialMasterName(kind)} master points to record 0x{record.Type:X4} instead of 0x{expectedRecordType:X4}.",
                    record.Offset);
                return null;
            }

            if (kind == LegacyPptSpecialMasterKind.Notes) ValidateNotesMasterAtom(record, options);
            var master = new LegacyPptSpecialMaster(kind, persistId) {
                ColorScheme = ReadColorScheme(record),
                RoundTripTheme = ReadRoundTripTheme(record,
                    $"{GetSpecialMasterName(kind)} master", options)
            };
            master.Background = ReadBackground(record, master.ColorScheme, options);
            ParseShapes(record, master.AddShape, $"{GetSpecialMasterName(kind)} master", options,
                master.ColorScheme, master.AddConnectorRule);
            return master;
        }

        private void ValidateNotesMasterAtom(LegacyPptRecord record,
            LegacyPptImportOptions options) {
            LegacyPptRecord? atom = record.Children.FirstOrDefault(child => child.Type == RecordNotesAtom);
            if (atom == null || atom.PayloadLength < 8) {
                AddDiagnostic("PPT-NOTES-MASTER-ATOM", LegacyPptDiagnosticSeverity.Warning,
                    "The notes master has no complete NotesAtom and remains preserve-only.",
                    atom?.Offset ?? record.Offset);
                return;
            }
            if (options.ReportUnsupportedContent
                && (atom.ReadUInt32(0) != 0 || atom.ReadUInt16(4) != 0)) {
                AddDiagnostic("PPT-NOTES-MASTER-ATOM-INVALID",
                    LegacyPptDiagnosticSeverity.Information,
                    "The notes master NotesAtom contains slide-specific fields; they have no master semantics and were ignored.",
                    atom.Offset);
            }
        }

        private static string GetSpecialMasterName(LegacyPptSpecialMasterKind kind) =>
            kind == LegacyPptSpecialMasterKind.Notes ? "notes" : "handout";
    }
}
