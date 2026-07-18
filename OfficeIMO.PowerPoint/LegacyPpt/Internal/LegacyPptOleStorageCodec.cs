using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Decodes the shared ExOleObjStg, ExControlStg, and VbaProjectStg record.</summary>
    internal static class LegacyPptOleStorageCodec {
        private const ushort RecordExternalOleObjectStorage = 0x1011;

        internal static bool TryDecode(LegacyPptPersistObject persistObject,
            LegacyPptImportOptions options,
            LegacyPptRecordTraversalBudget recordBudget,
            LegacyPptDecodedStorageBudget decodedStorageBudget,
            out byte[] storageBytes,
            out bool wasCompressed, out string? reason) {
            if (persistObject == null) {
                throw new ArgumentNullException(nameof(persistObject));
            }
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (recordBudget == null) {
                throw new ArgumentNullException(nameof(recordBudget));
            }
            if (decodedStorageBudget == null) {
                throw new ArgumentNullException(nameof(decodedStorageBudget));
            }
            storageBytes = Array.Empty<byte>();
            wasCompressed = false;
            reason = null;
            if (persistObject.RecordType != RecordExternalOleObjectStorage) {
                reason = $"Persist object {persistObject.PersistId} is record "
                    + $"0x{persistObject.RecordType:X4}, not ExOleObjStg.";
                return false;
            }
            try {
                LegacyPptRecord storage = LegacyPptRecordReader.ReadSingle(
                    persistObject.RecordBytes, 0, options, recordBudget);
                if (storage.Version != 0
                    || storage.Type != RecordExternalOleObjectStorage
                    || (storage.Instance != 0 && storage.Instance != 1)) {
                    reason = "The storage has an unsupported record header.";
                    return false;
                }
                wasCompressed = storage.Instance == 1;
                if (!wasCompressed) {
                    decodedStorageBudget.Consume(storage.PayloadLength);
                    storageBytes = new byte[storage.PayloadLength];
                    Buffer.BlockCopy(persistObject.RecordBytes, 8,
                        storageBytes, 0, storageBytes.Length);
                    return true;
                }
                if (storage.PayloadLength < 4) {
                    reason = "The compressed storage is truncated.";
                    return false;
                }
                uint decompressedSize = storage.ReadUInt32(0);
                if (decompressedSize > options.MaxInputBytes) {
                    reason = $"The decompressed storage exceeds "
                        + $"{options.MaxInputBytes} bytes.";
                    return false;
                }
                decodedStorageBudget.Consume(checked((int)decompressedSize));
                var compressedBytes = new byte[storage.PayloadLength - 4];
                Buffer.BlockCopy(persistObject.RecordBytes, 12,
                    compressedBytes, 0, compressedBytes.Length);
                storageBytes = OfficeZlibCodec.Decompress(compressedBytes,
                    options.MaxInputBytes, checked((int)decompressedSize));
                return true;
            } catch (Exception exception) when (exception is InvalidDataException
                                                or NotSupportedException
                                                or OverflowException) {
                storageBytes = Array.Empty<byte>();
                reason = exception.Message;
                return false;
            }
        }
    }
}
