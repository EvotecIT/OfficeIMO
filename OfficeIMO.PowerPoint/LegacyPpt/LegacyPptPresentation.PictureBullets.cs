using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordBlipCollection9 = 0x07F8;
        private const ushort RecordBlipEntity9Atom = 0x07F9;

        private void ParsePictureBullets(LegacyPptRecord document,
            LegacyPptImportOptions options) {
            LegacyPptRecord[] infoLists = document.Children.Where(record =>
                record.Type == RecordDocumentInfoList).ToArray();
            if (infoLists.Length != 1) return;
            LegacyPptRecord[] ppt9Tags = infoLists[0].Children
                .Where(record => record.Type == RecordProgTags)
                .SelectMany(record => record.Children)
                .Where(IsPpt9BinaryTag).ToArray();
            if (ppt9Tags.Length > 1) {
                AddDiagnostic("PPT-PICTURE-BULLET-TAG-DUPLICATE",
                    LegacyPptDiagnosticSeverity.Warning,
                    "The document has multiple PPT9 tags; picture bullets remain preserve-only.",
                    ppt9Tags[1].Offset);
                return;
            }
            if (ppt9Tags.Length == 0) return;
            LegacyPptRecord[] blobs = ppt9Tags[0].Children.Where(record =>
                record.Type == RecordBinaryTagDataBlob).ToArray();
            if (blobs.Length != 1 || blobs[0].Version != 0
                || blobs[0].Instance != 0) {
                AddDiagnostic("PPT-PICTURE-BULLET-DATA",
                    LegacyPptDiagnosticSeverity.Warning,
                    "The document PPT9 tag has no unique valid data blob; picture bullets remain preserve-only.",
                    ppt9Tags[0].Offset);
                return;
            }
            IReadOnlyList<LegacyPptRecord> records;
            try {
                records = LegacyPptRecordReader.ReadSequence(
                    blobs[0].CopyRecordBytes(), 8, blobs[0].PayloadLength,
                    options, _recordBudget);
            } catch (Exception exception) when (exception
                is InvalidDataException or OverflowException
                    or ArgumentOutOfRangeException) {
                AddDiagnostic("PPT-PICTURE-BULLET-DATA-TRUNCATED",
                    LegacyPptDiagnosticSeverity.Warning,
                    "The document PPT9 data is malformed; picture bullets remain preserve-only.",
                    blobs[0].Offset);
                return;
            }
            LegacyPptRecord[] collections = records.Where(record =>
                record.Type == RecordBlipCollection9).ToArray();
            if (collections.Length > 1) {
                AddDiagnostic("PPT-PICTURE-BULLET-COLLECTION-DUPLICATE",
                    LegacyPptDiagnosticSeverity.Warning,
                    "The document has multiple PPT9 picture-bullet collections; they remain preserve-only.",
                    collections[1].Offset);
                return;
            }
            if (collections.Length == 0) return;
            LegacyPptRecord collection = collections[0];
            if (collection.Version != 0x0F || collection.Instance != 0) {
                AddDiagnostic("PPT-PICTURE-BULLET-COLLECTION",
                    LegacyPptDiagnosticSeverity.Warning,
                    "The PPT9 picture-bullet collection has an invalid header and remains preserve-only.",
                    collection.Offset);
                return;
            }
            foreach (LegacyPptRecord entity in collection.Children) {
                if (!TryReadPictureBullet(entity, options,
                        out LegacyPptPictureBullet? pictureBullet)
                    || pictureBullet == null) continue;
                if (_pictureBulletsByIndex.ContainsKey(pictureBullet.Index)) {
                    AddDiagnostic("PPT-PICTURE-BULLET-INDEX-DUPLICATE",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"Picture-bullet index {pictureBullet.Index} is duplicated and remains preserve-only.",
                        entity.Offset);
                    _pictureBulletsByIndex.Remove(pictureBullet.Index);
                    _pictureBullets.RemoveAll(item => item.Index
                        == pictureBullet.Index);
                    continue;
                }
                _pictureBulletsByIndex.Add(pictureBullet.Index,
                    pictureBullet);
                _pictureBullets.Add(pictureBullet);
            }
            _pictureBullets.Sort((left, right) => left.Index
                .CompareTo(right.Index));
        }

        private bool TryReadPictureBullet(LegacyPptRecord entity,
            LegacyPptImportOptions options,
            out LegacyPptPictureBullet? pictureBullet) {
            pictureBullet = null;
            if (entity.Type != RecordBlipEntity9Atom
                || entity.Version != 0 || entity.Instance > 0x080
                || entity.PayloadLength < 10) {
                if (options.ReportUnsupportedContent) {
                    AddDiagnostic("PPT-PICTURE-BULLET-ENTITY",
                        LegacyPptDiagnosticSeverity.Warning,
                        "A PPT9 picture-bullet entity has an invalid header or payload and remains preserve-only.",
                        entity.Offset);
                }
                return false;
            }
            byte[] bytes = entity.CopyRecordBytes();
            byte preferredType = bytes[8];
            int remainingDecodedBytes = _decodedStorageBudget
                .RemainingAllocationBytes;
            int perImageLimit = Math.Min(options.MaxInputBytes,
                remainingDecodedBytes);
            if (preferredType is not 0x02 and not 0x03 and not 0x05
                and not 0x06
                || !OfficeArtBlipStoreEntryReader.TryReadBlipStoreFileBlock(bytes,
                    10, entity.PayloadLength - 2,
                    out OfficeArtBlipStoreEntryReader
                        .OfficeArtBlipRecordData blip, perImageLimit)) {
                AddDiagnostic("PPT-PICTURE-BULLET-BLIP",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Picture-bullet index {entity.Instance} has an unsupported or malformed OfficeArt BLIP.",
                    entity.Offset);
                return false;
            }
            if (blip.WasImageRejectedBySizeLimit
                && remainingDecodedBytes <= options.MaxInputBytes) {
                _decodedStorageBudget.RejectAllocation();
            }
            byte[] imageBytes = blip.ImageBytes;
            _decodedStorageBudget.Consume(imageBytes.Length);
            pictureBullet = new LegacyPptPictureBullet(entity.Instance,
                preferredType, bytes[9], blip.RecordVersion,
                blip.RecordInstance, blip.RecordType, blip.PayloadLength,
                blip.PayloadAvailableLength, blip.PayloadSha256,
                blip.ContentType, imageBytes,
                blip.IsPayloadTruncated);
            if (!pictureBullet.HasImportableImage) {
                if (pictureBullet.IsPayloadTruncated) {
                    AddDiagnostic(
                        "PPT-PICTURE-BULLET-BLIP-TRUNCATED",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"Picture-bullet index {entity.Instance} has a truncated image payload and remains preserve-only.",
                        entity.Offset);
                } else {
                    AddDiagnostic("PPT-PICTURE-BULLET-IMAGE",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"Picture-bullet index {entity.Instance} has no bounded importable image payload and remains preserve-only.",
                        entity.Offset);
                }
            }
            return true;
        }
    }
}
