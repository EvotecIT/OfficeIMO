using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordDocInfoList = 0x07D0;
        private const ushort RecordVbaInfo = 0x03FF;
        private const ushort RecordVbaInfoAtom = 0x0400;
        private const ushort RecordExternalOleObjectStorage = 0x1011;

        private void ParseVbaProject(LegacyPptRecord document,
            LegacyPptPackage package, LegacyPptImportOptions options) {
            LegacyPptRecord? docInfo = document.Children.FirstOrDefault(
                child => child.Type == RecordDocInfoList);
            LegacyPptRecord? vbaInfo = docInfo?.Children.FirstOrDefault(
                child => child.Type == RecordVbaInfo);
            LegacyPptRecord? atom = vbaInfo?.Children.FirstOrDefault(
                child => child.Type == RecordVbaInfoAtom);
            if (vbaInfo == null) return;
            if (vbaInfo.Version != 0x0F || vbaInfo.Instance != 1
                || atom == null || atom.Version != 2 || atom.Instance != 0
                || atom.PayloadLength != 12) {
                AddDiagnostic("PPT-VBA-INFO-MALFORMED",
                    LegacyPptDiagnosticSeverity.Warning,
                    "The VBA information container is malformed and remains preserve-only.",
                    vbaInfo.Offset);
                return;
            }

            uint persistId = atom.ReadUInt32(0);
            uint hasMacros = atom.ReadUInt32(4);
            uint version = atom.ReadUInt32(8);
            if (hasMacros == 0) return;
            if (hasMacros != 1 || version != 2
                || !package.PersistObjects.TryGetValue(persistId,
                    out LegacyPptPersistObject? persistObject)
                || persistObject.RecordType != RecordExternalOleObjectStorage) {
                AddDiagnostic("PPT-VBA-STORAGE-MISSING",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"The VBA information references invalid persist object {persistId}; it remains preserve-only.",
                    atom.Offset);
                return;
            }

            try {
                if (!LegacyPptOleStorageCodec.TryDecode(persistObject,
                        options, _recordBudget, _decodedStorageBudget,
                        out byte[] projectBytes,
                        out bool compressed,
                        out string? storageReason)) {
                    throw new InvalidDataException(storageReason
                        ?? "The VBA project storage cannot be decoded.");
                }

                if (!LegacyPptVbaProjectCodec.IsValidProject(projectBytes,
                        options,
                        out string? error)) {
                    throw new InvalidDataException(error
                        ?? "The referenced VBA project is not a valid VBA compound storage.");
                }
                VbaProject = new LegacyPptVbaProject(persistId, compressed,
                    projectBytes);
            } catch (Exception exception) when (exception is InvalidDataException
                                                or NotSupportedException) {
                AddDiagnostic("PPT-VBA-STORAGE-MALFORMED",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"The VBA project storage remains preserve-only: {exception.Message}",
                    unchecked((long)persistObject.StreamOffset));
            }
        }
    }
}
