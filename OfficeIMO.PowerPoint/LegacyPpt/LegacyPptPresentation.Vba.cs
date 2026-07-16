using OfficeIMO.Drawing.Internal;
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
                LegacyPptRecord storage = LegacyPptRecordReader.ReadSingle(
                    persistObject.RecordBytes, 0, options);
                if (storage.Version != 0 || storage.Type != RecordExternalOleObjectStorage
                    || (storage.Instance != 0 && storage.Instance != 1)) {
                    throw new InvalidDataException(
                        "The VBA persist object has an unsupported record header.");
                }

                bool compressed = storage.Instance == 1;
                byte[] projectBytes;
                if (compressed) {
                    if (storage.PayloadLength < 4) {
                        throw new InvalidDataException(
                            "The compressed VBA persist object is truncated.");
                    }
                    uint decompressedSize = storage.ReadUInt32(0);
                    if (decompressedSize > options.MaxInputBytes) {
                        throw new InvalidDataException(
                            $"The VBA project exceeds {options.MaxInputBytes} bytes.");
                    }
                    var compressedBytes = new byte[storage.PayloadLength - 4];
                    Buffer.BlockCopy(persistObject.RecordBytes, 12,
                        compressedBytes, 0, compressedBytes.Length);
                    projectBytes = OfficeZlibCodec.Decompress(compressedBytes,
                        options.MaxInputBytes, checked((int)decompressedSize));
                } else {
                    projectBytes = new byte[storage.PayloadLength];
                    Buffer.BlockCopy(persistObject.RecordBytes, 8,
                        projectBytes, 0, projectBytes.Length);
                }

                if (!LegacyPptVbaProjectCodec.IsValidProject(projectBytes,
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
