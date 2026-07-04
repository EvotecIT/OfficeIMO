using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffDrawingMetadataReader {
        internal static bool TryRead(
            BiffRecord record,
            string? sheetName,
            List<LegacyXlsDrawingRecord> records) {
            if (TryRead(record, sheetName, out LegacyXlsDrawingRecord? drawingRecord)) {
                records.Add(drawingRecord!);
                return true;
            }

            return false;
        }

        internal static bool TryRead(
            BiffRecord record,
            string? sheetName,
            out LegacyXlsDrawingRecord? drawingRecord) {
            drawingRecord = null;
            if (!BiffUnsupportedRecordDiagnostics.IsDrawingRecord(record.Type)) {
                return false;
            }

            TryReadObjectCommonData(record, out ushort? objectType, out ushort? objectId, out ushort? objectFlags);
            IReadOnlyList<LegacyXlsDrawingObjectSubRecord> objectSubRecords = ReadObjectSubRecords(record);
            LegacyXlsDrawingFutureRecordHeader? futureRecordHeader = TryReadFutureRecordHeader(record);
            LegacyXlsDrawingTextObject? textObject = TryReadTextObject(record);
            LegacyXlsHeaderFooterPicture? headerFooterPicture = TryReadHeaderFooterPicture(record);
            TryReadEscherHeader(record, out ushort? escherRecordType, out ushort? escherRecordInstance, out byte? escherRecordVersion, out uint? escherPayloadLength);
            ReadOfficeArtMetadata(
                record,
                out IReadOnlyList<LegacyXlsDrawingBlipStoreEntry> blipStoreEntries,
                out IReadOnlyList<LegacyXlsDrawingShape> shapeEntries,
                out IReadOnlyList<LegacyXlsDrawingAnchor> anchorEntries,
                out IReadOnlyList<LegacyXlsDrawingChildAnchor> childAnchorEntries,
                out IReadOnlyList<LegacyXlsDrawingOfficeArtRecord> officeArtRecords,
                out IReadOnlyList<LegacyXlsDrawingGroupBlock> drawingGroupBlocks,
                out IReadOnlyList<LegacyXlsDrawingGroupInfo> drawingGroupInfos,
                out IReadOnlyList<LegacyXlsDrawingShapeProperty> shapeProperties,
                out bool officeArtPayloadFullyTraversed);
            drawingRecord = new LegacyXlsDrawingRecord(
                GetKind(record.Type),
                BiffUnsupportedRecordDiagnostics.GetBiffRecordName(record.Type),
                sheetName,
                record.Offset,
                record.Type,
                record.Payload.Length,
                objectType,
                objectId,
                escherRecordType,
                escherRecordInstance,
                escherRecordVersion,
                escherPayloadLength,
                objectFlags: objectFlags,
                blipStoreEntries: blipStoreEntries,
                shapeEntries: shapeEntries,
                anchorEntries: anchorEntries,
                childAnchorEntries: childAnchorEntries,
                officeArtRecords: officeArtRecords,
                drawingGroupBlocks: drawingGroupBlocks,
                drawingGroupInfos: drawingGroupInfos,
                shapeProperties: shapeProperties,
                objectSubRecords: objectSubRecords,
                futureRecordHeader: futureRecordHeader,
                textObject: textObject,
                headerFooterPicture: headerFooterPicture,
                officeArtPayloadFullyTraversed: officeArtPayloadFullyTraversed);
            return true;
        }

        internal static bool TryReadClientAnchors(
            BiffRecord record,
            out IReadOnlyList<LegacyXlsDrawingAnchor> anchors) {
            if (record.Type != (ushort)BiffRecordType.Drawing) {
                anchors = Array.Empty<LegacyXlsDrawingAnchor>();
                return false;
            }

            ReadOfficeArtMetadata(
                record,
                out _,
                out _,
                out anchors,
                out _,
                out _,
                out _,
                out _,
                out _,
                out _);
            return anchors.Count > 0;
        }

        private static void ReadOfficeArtMetadata(
            BiffRecord record,
            out IReadOnlyList<LegacyXlsDrawingBlipStoreEntry> blipStoreEntries,
            out IReadOnlyList<LegacyXlsDrawingShape> shapeEntries,
            out IReadOnlyList<LegacyXlsDrawingAnchor> anchorEntries,
            out IReadOnlyList<LegacyXlsDrawingChildAnchor> childAnchorEntries,
            out IReadOnlyList<LegacyXlsDrawingOfficeArtRecord> officeArtRecords,
            out IReadOnlyList<LegacyXlsDrawingGroupBlock> drawingGroupBlocks,
            out IReadOnlyList<LegacyXlsDrawingGroupInfo> drawingGroupInfos,
            out IReadOnlyList<LegacyXlsDrawingShapeProperty> shapeProperties,
            out bool officeArtPayloadFullyTraversed) {
            if (!TryGetOfficeArtPayloadRange(record, out int officeArtOffset, out int officeArtLength)) {
                blipStoreEntries = Array.Empty<LegacyXlsDrawingBlipStoreEntry>();
                shapeEntries = Array.Empty<LegacyXlsDrawingShape>();
                anchorEntries = Array.Empty<LegacyXlsDrawingAnchor>();
                childAnchorEntries = Array.Empty<LegacyXlsDrawingChildAnchor>();
                officeArtRecords = Array.Empty<LegacyXlsDrawingOfficeArtRecord>();
                drawingGroupBlocks = Array.Empty<LegacyXlsDrawingGroupBlock>();
                drawingGroupInfos = Array.Empty<LegacyXlsDrawingGroupInfo>();
                shapeProperties = Array.Empty<LegacyXlsDrawingShapeProperty>();
                officeArtPayloadFullyTraversed = false;
                return;
            }

            ReadOfficeArtPayload(
                record.Payload,
                officeArtOffset,
                officeArtLength,
                out blipStoreEntries,
                out shapeEntries,
                out anchorEntries,
                out childAnchorEntries,
                out officeArtRecords,
                out drawingGroupBlocks,
                out drawingGroupInfos,
                out shapeProperties,
                out officeArtPayloadFullyTraversed);
        }

        internal static void ReadOfficeArtPayload(
            byte[] payload,
            out IReadOnlyList<LegacyXlsDrawingBlipStoreEntry> blipStoreEntries,
            out IReadOnlyList<LegacyXlsDrawingShape> shapeEntries,
            out IReadOnlyList<LegacyXlsDrawingAnchor> anchorEntries,
            out IReadOnlyList<LegacyXlsDrawingChildAnchor> childAnchorEntries,
            out IReadOnlyList<LegacyXlsDrawingOfficeArtRecord> officeArtRecords,
            out IReadOnlyList<LegacyXlsDrawingGroupBlock> drawingGroupBlocks,
            out IReadOnlyList<LegacyXlsDrawingGroupInfo> drawingGroupInfos,
            out IReadOnlyList<LegacyXlsDrawingShapeProperty> shapeProperties,
            out bool fullyTraversed) {
            if (payload == null) {
                throw new ArgumentNullException(nameof(payload));
            }

            var blips = new List<LegacyXlsDrawingBlipStoreEntry>();
            var shapes = new List<LegacyXlsDrawingShape>();
            var anchors = new List<LegacyXlsDrawingAnchor>();
            var childAnchors = new List<LegacyXlsDrawingChildAnchor>();
            var records = new List<LegacyXlsDrawingOfficeArtRecord>();
            var groupBlocks = new List<LegacyXlsDrawingGroupBlock>();
            var groupInfos = new List<LegacyXlsDrawingGroupInfo>();
            var properties = new List<LegacyXlsDrawingShapeProperty>();
            fullyTraversed = TryReadOfficeArtRecords(payload, 0, payload.Length, records, blips, shapes, anchors, childAnchors, groupBlocks, groupInfos, properties, depth: 0);
            blipStoreEntries = blips;
            shapeEntries = shapes;
            anchorEntries = anchors;
            childAnchorEntries = childAnchors;
            officeArtRecords = records;
            drawingGroupBlocks = groupBlocks;
            drawingGroupInfos = groupInfos;
            shapeProperties = properties;
        }

        private static void ReadOfficeArtPayload(
            byte[] payload,
            int offset,
            int length,
            out IReadOnlyList<LegacyXlsDrawingBlipStoreEntry> blipStoreEntries,
            out IReadOnlyList<LegacyXlsDrawingShape> shapeEntries,
            out IReadOnlyList<LegacyXlsDrawingAnchor> anchorEntries,
            out IReadOnlyList<LegacyXlsDrawingChildAnchor> childAnchorEntries,
            out IReadOnlyList<LegacyXlsDrawingOfficeArtRecord> officeArtRecords,
            out IReadOnlyList<LegacyXlsDrawingGroupBlock> drawingGroupBlocks,
            out IReadOnlyList<LegacyXlsDrawingGroupInfo> drawingGroupInfos,
            out IReadOnlyList<LegacyXlsDrawingShapeProperty> shapeProperties,
            out bool fullyTraversed) {
            if (payload == null) {
                throw new ArgumentNullException(nameof(payload));
            }

            var blips = new List<LegacyXlsDrawingBlipStoreEntry>();
            var shapes = new List<LegacyXlsDrawingShape>();
            var anchors = new List<LegacyXlsDrawingAnchor>();
            var childAnchors = new List<LegacyXlsDrawingChildAnchor>();
            var records = new List<LegacyXlsDrawingOfficeArtRecord>();
            var groupBlocks = new List<LegacyXlsDrawingGroupBlock>();
            var groupInfos = new List<LegacyXlsDrawingGroupInfo>();
            var properties = new List<LegacyXlsDrawingShapeProperty>();
            int endOffset = offset + length;
            fullyTraversed = offset >= 0
                && length >= 0
                && endOffset <= payload.Length
                && TryReadOfficeArtRecords(payload, offset, endOffset, records, blips, shapes, anchors, childAnchors, groupBlocks, groupInfos, properties, depth: 0);
            blipStoreEntries = blips;
            shapeEntries = shapes;
            anchorEntries = anchors;
            childAnchorEntries = childAnchors;
            officeArtRecords = records;
            drawingGroupBlocks = groupBlocks;
            drawingGroupInfos = groupInfos;
            shapeProperties = properties;
        }

        private static bool TryReadOfficeArtRecords(
            byte[] payload,
            int startOffset,
            int endOffset,
            List<LegacyXlsDrawingOfficeArtRecord> officeArtRecords,
            List<LegacyXlsDrawingBlipStoreEntry> blipStoreEntries,
            List<LegacyXlsDrawingShape> shapeEntries,
            List<LegacyXlsDrawingAnchor> anchorEntries,
            List<LegacyXlsDrawingChildAnchor> childAnchorEntries,
            List<LegacyXlsDrawingGroupBlock> drawingGroupBlocks,
            List<LegacyXlsDrawingGroupInfo> drawingGroupInfos,
            List<LegacyXlsDrawingShapeProperty> shapeProperties,
            int depth) {
            if (depth > 8) {
                return false;
            }

            int offset = startOffset;
            bool fullyTraversed = true;
            while (offset + 8 <= endOffset) {
                ushort options = BiffRecordReader.ReadUInt16(payload, offset);
                ushort recordType = BiffRecordReader.ReadUInt16(payload, offset + 2);
                uint recordLength = BiffRecordReader.ReadUInt32(payload, offset + 4);
                byte version = checked((byte)(options & 0x000f));
                ushort instance = checked((ushort)(options >> 4));
                int contentStart = offset + 8;
                if (recordLength > int.MaxValue || contentStart > endOffset) {
                    return false;
                }

                if (contentStart + (int)recordLength > endOffset) {
                    if (version == 0x0f) {
                        officeArtRecords.Add(new LegacyXlsDrawingOfficeArtRecord(recordType, instance, version, recordLength, depth));
                        TryReadOfficeArtRecords(payload, contentStart, endOffset, officeArtRecords, blipStoreEntries, shapeEntries, anchorEntries, childAnchorEntries, drawingGroupBlocks, drawingGroupInfos, shapeProperties, depth + 1);
                    }

                    return false;
                }

                int contentEnd = contentStart + (int)recordLength;
                officeArtRecords.Add(new LegacyXlsDrawingOfficeArtRecord(recordType, instance, version, recordLength, depth));

                if (recordType == 0xF007 && TryReadBlipStoreEntry(payload, contentStart, contentEnd, instance, out LegacyXlsDrawingBlipStoreEntry? blipEntry)) {
                    blipStoreEntries.Add(blipEntry!);
                } else if (recordType == 0xF006 && TryReadDrawingGroupBlock(payload, contentStart, contentEnd, out LegacyXlsDrawingGroupBlock? drawingGroupBlock)) {
                    drawingGroupBlocks.Add(drawingGroupBlock!);
                } else if (recordType == 0xF008 && TryReadDrawingGroupInfo(payload, contentStart, contentEnd, instance, out LegacyXlsDrawingGroupInfo? drawingGroupInfo)) {
                    drawingGroupInfos.Add(drawingGroupInfo!);
                } else if (recordType == 0xF00A && TryReadShape(payload, contentStart, contentEnd, instance, out LegacyXlsDrawingShape? shapeEntry)) {
                    shapeEntries.Add(shapeEntry!);
                } else if (recordType == 0xF010 && TryReadClientAnchor(payload, contentStart, contentEnd, out LegacyXlsDrawingAnchor? anchorEntry)) {
                    anchorEntries.Add(anchorEntry!);
                } else if (recordType == 0xF00F && TryReadChildAnchor(payload, contentStart, contentEnd, out LegacyXlsDrawingChildAnchor? childAnchorEntry)) {
                    childAnchorEntries.Add(childAnchorEntry!);
                } else if (recordType == 0xF00B) {
                    TryReadShapeProperties(payload, contentStart, contentEnd, instance, shapeProperties);
                }

                if (version == 0x0f) {
                    fullyTraversed &= TryReadOfficeArtRecords(payload, contentStart, contentEnd, officeArtRecords, blipStoreEntries, shapeEntries, anchorEntries, childAnchorEntries, drawingGroupBlocks, drawingGroupInfos, shapeProperties, depth + 1);
                }

                offset = contentEnd;
            }

            return fullyTraversed && offset == endOffset;
        }

        private static bool TryReadDrawingGroupBlock(
            byte[] payload,
            int contentStart,
            int contentEnd,
            out LegacyXlsDrawingGroupBlock? block) {
            block = null;
            if (contentStart < 0 || contentStart + 16 > contentEnd) {
                return false;
            }

            uint identifierClusterCount = BiffRecordReader.ReadUInt32(payload, contentStart + 4);
            int declaredClusterCount = identifierClusterCount > 0 && identifierClusterCount <= int.MaxValue
                ? checked((int)identifierClusterCount - 1)
                : 0;
            int availableClusterCount = Math.Max(0, (contentEnd - (contentStart + 16)) / 8);
            int clusterCount = Math.Min(declaredClusterCount, availableClusterCount);
            var clusters = new List<LegacyXlsDrawingIdentifierCluster>(clusterCount);
            int clusterOffset = contentStart + 16;
            for (int i = 0; i < clusterCount; i++) {
                clusters.Add(new LegacyXlsDrawingIdentifierCluster(
                    BiffRecordReader.ReadUInt32(payload, clusterOffset),
                    BiffRecordReader.ReadUInt32(payload, clusterOffset + 4)));
                clusterOffset += 8;
            }

            block = new LegacyXlsDrawingGroupBlock(
                BiffRecordReader.ReadUInt32(payload, contentStart),
                identifierClusterCount,
                BiffRecordReader.ReadUInt32(payload, contentStart + 8),
                BiffRecordReader.ReadUInt32(payload, contentStart + 12),
                clusters);
            return true;
        }

        private static bool TryReadDrawingGroupInfo(
            byte[] payload,
            int contentStart,
            int contentEnd,
            ushort drawingId,
            out LegacyXlsDrawingGroupInfo? info) {
            info = null;
            if (contentStart < 0 || contentStart + 8 > contentEnd) {
                return false;
            }

            info = new LegacyXlsDrawingGroupInfo(
                drawingId,
                BiffRecordReader.ReadUInt32(payload, contentStart),
                BiffRecordReader.ReadUInt32(payload, contentStart + 4));
            return true;
        }

        private static bool TryReadBlipStoreEntry(
            byte[] payload,
            int contentStart,
            int contentEnd,
            ushort recordInstance,
            out LegacyXlsDrawingBlipStoreEntry? entry) {
            entry = null;
            if (contentStart < 0 || contentStart + 36 > contentEnd) {
                return false;
            }

            byte win32BlipType = payload[contentStart];
            byte macOsBlipType = payload[contentStart + 1];
            string uidHex = ToHexString(payload, contentStart + 2, 16);
            uint sizeBytes = BiffRecordReader.ReadUInt32(payload, contentStart + 20);
            uint referenceCount = BiffRecordReader.ReadUInt32(payload, contentStart + 24);
            byte nameByteCount = payload[contentStart + 33];
            int embeddedOffset = contentStart + 36 + nameByteCount;
            ushort? embeddedRecordType = null;
            uint? embeddedPayloadLength = null;
            int? embeddedPayloadAvailableLength = null;
            string? embeddedPayloadSha256 = null;
            byte[]? embeddedPayloadBytes = null;
            if (embeddedOffset + 8 <= contentEnd) {
                embeddedRecordType = BiffRecordReader.ReadUInt16(payload, embeddedOffset + 2);
                embeddedPayloadLength = BiffRecordReader.ReadUInt32(payload, embeddedOffset + 4);
                int embeddedPayloadOffset = embeddedOffset + 8;
                int declaredEmbeddedPayloadLength = embeddedPayloadLength > int.MaxValue
                    ? int.MaxValue
                    : (int)embeddedPayloadLength.Value;
                embeddedPayloadAvailableLength = Math.Min(Math.Max(0, contentEnd - embeddedPayloadOffset), declaredEmbeddedPayloadLength);
                if (embeddedPayloadAvailableLength > 0) {
                    embeddedPayloadSha256 = ComputeSha256(payload, embeddedPayloadOffset, embeddedPayloadAvailableLength.Value);
                    embeddedPayloadBytes = CopyEmbeddedBlipPayload(payload, embeddedPayloadOffset, embeddedPayloadAvailableLength.Value, embeddedRecordType);
                }
            }

            entry = new LegacyXlsDrawingBlipStoreEntry(
                recordInstance,
                win32BlipType,
                macOsBlipType,
                uidHex,
                sizeBytes,
                referenceCount,
                embeddedRecordType,
                embeddedPayloadLength,
                embeddedPayloadAvailableLength,
                embeddedPayloadSha256,
                embeddedPayloadBytes);
            return true;
        }

        private static byte[] CopyEmbeddedBlipPayload(byte[] payload, int offset, int count, ushort? embeddedRecordType) {
            int imageOffset = FindEmbeddedImageOffset(payload, offset, count, embeddedRecordType);
            int imageLength = count - (imageOffset - offset);
            if (imageLength <= 0) {
                return Array.Empty<byte>();
            }

            var bytes = new byte[imageLength];
            Buffer.BlockCopy(payload, imageOffset, bytes, 0, imageLength);
            return bytes;
        }

        private static int FindEmbeddedImageOffset(byte[] payload, int offset, int count, ushort? embeddedRecordType) {
            if (!embeddedRecordType.HasValue) {
                return offset;
            }

            int maxScan = Math.Min(count, 32);
            switch (embeddedRecordType.Value) {
                case 0xF01D:
                case 0xF02A:
                    for (int i = 0; i + 1 < maxScan; i++) {
                        if (payload[offset + i] == 0xff && payload[offset + i + 1] == 0xd8) {
                            return offset + i;
                        }
                    }

                    break;
                case 0xF01E:
                    for (int i = 0; i + 3 < maxScan; i++) {
                        if (payload[offset + i] == 0x89
                            && payload[offset + i + 1] == 0x50
                            && payload[offset + i + 2] == 0x4e
                            && payload[offset + i + 3] == 0x47) {
                            return offset + i;
                        }
                    }

                    break;
                case 0xF01F:
                    for (int i = 0; i + 1 < maxScan; i++) {
                        if (payload[offset + i] == 0x42 && payload[offset + i + 1] == 0x4d) {
                            return offset + i;
                        }
                    }

                    break;
                case 0xF029:
                    for (int i = 0; i + 3 < maxScan; i++) {
                        if ((payload[offset + i] == 0x49 && payload[offset + i + 1] == 0x49 && payload[offset + i + 2] == 0x2a && payload[offset + i + 3] == 0x00)
                            || (payload[offset + i] == 0x4d && payload[offset + i + 1] == 0x4d && payload[offset + i + 2] == 0x00 && payload[offset + i + 3] == 0x2a)) {
                            return offset + i;
                        }
                    }

                    break;
            }

            return offset;
        }

        private static string ComputeSha256(byte[] payload, int offset, int count) {
            using SHA256 sha256 = SHA256.Create();
            byte[] hash = sha256.ComputeHash(payload, offset, count);
            return ToHexString(hash, 0, hash.Length);
        }

        private static string ToHexString(byte[] payload, int offset, int count) {
            var builder = new StringBuilder(count * 2);
            for (int i = 0; i < count; i++) {
                builder.Append(payload[offset + i].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
            }

            return builder.ToString();
        }

        private static bool TryReadShape(
            byte[] payload,
            int contentStart,
            int contentEnd,
            ushort recordInstance,
            out LegacyXlsDrawingShape? shape) {
            shape = null;
            if (contentStart < 0 || contentStart + 8 > contentEnd) {
                return false;
            }

            shape = new LegacyXlsDrawingShape(
                recordInstance,
                BiffRecordReader.ReadUInt32(payload, contentStart),
                BiffRecordReader.ReadUInt32(payload, contentStart + 4));
            return true;
        }

        private static bool TryReadClientAnchor(
            byte[] payload,
            int contentStart,
            int contentEnd,
            out LegacyXlsDrawingAnchor? anchor) {
            anchor = null;
            if (contentStart < 0 || contentStart + 18 > contentEnd) {
                return false;
            }

            anchor = new LegacyXlsDrawingAnchor(
                BiffRecordReader.ReadUInt16(payload, contentStart),
                BiffRecordReader.ReadUInt16(payload, contentStart + 2),
                BiffRecordReader.ReadUInt16(payload, contentStart + 4),
                BiffRecordReader.ReadUInt16(payload, contentStart + 6),
                BiffRecordReader.ReadUInt16(payload, contentStart + 8),
                BiffRecordReader.ReadUInt16(payload, contentStart + 10),
                BiffRecordReader.ReadUInt16(payload, contentStart + 12),
                BiffRecordReader.ReadUInt16(payload, contentStart + 14),
                BiffRecordReader.ReadUInt16(payload, contentStart + 16));
            return true;
        }

        private static bool TryReadChildAnchor(
            byte[] payload,
            int contentStart,
            int contentEnd,
            out LegacyXlsDrawingChildAnchor? anchor) {
            anchor = null;
            if (contentStart < 0 || contentStart + 16 > contentEnd) {
                return false;
            }

            anchor = new LegacyXlsDrawingChildAnchor(
                BiffRecordReader.ReadInt32(payload, contentStart),
                BiffRecordReader.ReadInt32(payload, contentStart + 4),
                BiffRecordReader.ReadInt32(payload, contentStart + 8),
                BiffRecordReader.ReadInt32(payload, contentStart + 12));
            return true;
        }

        private static void TryReadShapeProperties(
            byte[] payload,
            int contentStart,
            int contentEnd,
            ushort propertyCount,
            List<LegacyXlsDrawingShapeProperty> properties) {
            int fixedLength = propertyCount * 6;
            if (propertyCount == 0 || contentStart < 0 || contentStart + fixedLength > contentEnd) {
                return;
            }

            int complexOffset = contentStart + fixedLength;
            for (int index = 0; index < propertyCount; index++) {
                int propertyOffset = contentStart + index * 6;
                ushort rawOperationId = BiffRecordReader.ReadUInt16(payload, propertyOffset);
                uint value = BiffRecordReader.ReadUInt32(payload, propertyOffset + 2);
                bool isComplex = (rawOperationId & 0x8000) != 0;
                int? availableComplexDataLength = null;
                string? complexText = null;
                if (isComplex) {
                    availableComplexDataLength = GetAvailableComplexDataLength(value, complexOffset, contentEnd);
                    ushort propertyId = checked((ushort)(rawOperationId & 0x3fff));
                    complexText = TryReadComplexTextProperty(payload, complexOffset, availableComplexDataLength.Value, propertyId);
                    complexOffset += availableComplexDataLength.Value;
                }

                properties.Add(new LegacyXlsDrawingShapeProperty(index, rawOperationId, value, availableComplexDataLength, complexText));
            }
        }

        private static string? TryReadComplexTextProperty(byte[] payload, int offset, int byteCount, ushort propertyId) {
            if (!IsComplexTextProperty(propertyId) || byteCount <= 0 || offset < 0 || offset > payload.Length || byteCount > payload.Length - offset) {
                return null;
            }

            int evenByteCount = byteCount - (byteCount % 2);
            if (evenByteCount <= 0) {
                return null;
            }

            string value = Encoding.Unicode.GetString(payload, offset, evenByteCount).TrimEnd('\0');
            return string.IsNullOrWhiteSpace(value) ? null : value;
        }

        private static bool IsComplexTextProperty(ushort propertyId) {
            return propertyId == 0x0380;
        }

        private static int GetAvailableComplexDataLength(uint declaredLength, int complexOffset, int contentEnd) {
            if (declaredLength > int.MaxValue || complexOffset >= contentEnd) {
                return 0;
            }

            int remaining = contentEnd - complexOffset;
            return Math.Min((int)declaredLength, remaining);
        }

        private static bool TryReadEscherHeader(
            BiffRecord record,
            out ushort? recordType,
            out ushort? recordInstance,
            out byte? recordVersion,
            out uint? payloadLength) {
            recordType = null;
            recordInstance = null;
            recordVersion = null;
            payloadLength = null;
            byte[] payload = record.Payload;
            if (!TryGetOfficeArtPayloadRange(record, out int officeArtOffset, out int officeArtLength)) {
                return false;
            }

            if (officeArtLength < 8) {
                return false;
            }

            ushort options = BiffRecordReader.ReadUInt16(payload, officeArtOffset);
            recordVersion = checked((byte)(options & 0x000f));
            recordInstance = checked((ushort)(options >> 4));
            recordType = BiffRecordReader.ReadUInt16(payload, officeArtOffset + 2);
            payloadLength = BiffRecordReader.ReadUInt32(payload, officeArtOffset + 4);
            return true;
        }

        private static bool TryGetOfficeArtPayloadRange(BiffRecord record, out int offset, out int length) {
            offset = 0;
            length = 0;
            if (record.Type == (ushort)BiffRecordType.DrawingGroup ||
                record.Type == (ushort)BiffRecordType.Drawing) {
                length = record.Payload.Length;
                return true;
            }

            if (record.Type == (ushort)BiffRecordType.HfPicture && record.Payload.Length >= 14) {
                offset = 14;
                length = record.Payload.Length - 14;
                return true;
            }

            return false;
        }

        private static bool TryReadObjectCommonData(
            BiffRecord record,
            out ushort? objectType,
            out ushort? objectId,
            out ushort? objectFlags) {
            objectType = null;
            objectId = null;
            objectFlags = null;
            byte[] payload = record.Payload;
            if (record.Type != (ushort)BiffRecordType.Obj || payload.Length < 8) {
                return false;
            }

            ushort subRecordType = BiffRecordReader.ReadUInt16(payload, 0);
            ushort subRecordLength = BiffRecordReader.ReadUInt16(payload, 2);
            if (subRecordType != 0x0015 || subRecordLength < 4 || payload.Length < 8) {
                return false;
            }

            objectType = BiffRecordReader.ReadUInt16(payload, 4);
            objectId = BiffRecordReader.ReadUInt16(payload, 6);
            if (subRecordLength >= 6 && payload.Length >= 10) {
                objectFlags = BiffRecordReader.ReadUInt16(payload, 8);
            }

            return true;
        }

        private static LegacyXlsDrawingFutureRecordHeader? TryReadFutureRecordHeader(BiffRecord record) {
            if (record.Type != (ushort)BiffRecordType.ShapePropsStream &&
                record.Type != (ushort)BiffRecordType.TextPropsStream &&
                record.Type != (ushort)BiffRecordType.RichTextStream) {
                return null;
            }

            byte[] payload = record.Payload;
            if (payload.Length < 4) {
                return null;
            }

            ushort wrappedRecordType = BiffRecordReader.ReadUInt16(payload, 0);
            ushort flags = BiffRecordReader.ReadUInt16(payload, 2);
            ushort? firstRow = null;
            ushort? lastRow = null;
            ushort? firstColumn = null;
            ushort? lastColumn = null;
            int headerLength = 4;
            if ((flags & 0x0001) != 0 && payload.Length >= 12) {
                firstRow = BiffRecordReader.ReadUInt16(payload, 4);
                lastRow = BiffRecordReader.ReadUInt16(payload, 6);
                firstColumn = BiffRecordReader.ReadUInt16(payload, 8);
                lastColumn = BiffRecordReader.ReadUInt16(payload, 10);
                headerLength = 12;
            }

            return new LegacyXlsDrawingFutureRecordHeader(
                wrappedRecordType,
                flags,
                firstRow,
                lastRow,
                firstColumn,
                lastColumn,
                Math.Max(0, payload.Length - headerLength));
        }

        private static LegacyXlsDrawingTextObject? TryReadTextObject(BiffRecord record) {
            if (record.Type != (ushort)BiffRecordType.Txo || record.Payload.Length < 16) {
                return null;
            }

            byte[] payload = record.Payload;
            ushort options = BiffRecordReader.ReadUInt16(payload, 0);
            ushort rotation = BiffRecordReader.ReadUInt16(payload, 2);
            ushort textCharacterCount = BiffRecordReader.ReadUInt16(payload, 10);
            ushort formattingRunByteCount = BiffRecordReader.ReadUInt16(payload, 12);
            ushort emptyFontIndex = BiffRecordReader.ReadUInt16(payload, 14);
            return new LegacyXlsDrawingTextObject(
                options,
                rotation,
                textCharacterCount,
                formattingRunByteCount,
                emptyFontIndex,
                payload.Length - 16);
        }

        private static LegacyXlsHeaderFooterPicture? TryReadHeaderFooterPicture(BiffRecord record) {
            if (record.Type != (ushort)BiffRecordType.HfPicture || record.Payload.Length < 14) {
                return null;
            }

            byte[] payload = record.Payload;
            return new LegacyXlsHeaderFooterPicture(
                BiffRecordReader.ReadUInt16(payload, 0),
                BiffRecordReader.ReadUInt16(payload, 2),
                payload[12],
                payload[13],
                payload.Length - 14);
        }

        private static IReadOnlyList<LegacyXlsDrawingObjectSubRecord> ReadObjectSubRecords(BiffRecord record) {
            byte[] payload = record.Payload;
            if (record.Type != (ushort)BiffRecordType.Obj || payload.Length < 4) {
                return Array.Empty<LegacyXlsDrawingObjectSubRecord>();
            }

            var records = new List<LegacyXlsDrawingObjectSubRecord>();
            int offset = 0;
            while (offset + 4 <= payload.Length) {
                ushort subRecordType = BiffRecordReader.ReadUInt16(payload, offset);
                ushort declaredLength = BiffRecordReader.ReadUInt16(payload, offset + 2);
                int dataOffset = offset + 4;
                int availableLength = Math.Min(declaredLength, payload.Length - dataOffset);
                if (availableLength < 0) {
                    availableLength = 0;
                }

                records.Add(new LegacyXlsDrawingObjectSubRecord(subRecordType, offset, declaredLength, availableLength));
                if (subRecordType == 0x0000 || availableLength < declaredLength) {
                    break;
                }

                int nextOffset = dataOffset + availableLength;
                if ((declaredLength & 0x0001) != 0 && nextOffset < payload.Length) {
                    nextOffset++;
                }

                if (nextOffset <= offset) {
                    break;
                }

                offset = nextOffset;
            }

            return records;
        }

        private static LegacyXlsDrawingRecordKind GetKind(ushort type) {
            if (type == (ushort)BiffRecordType.DrawingGroup) {
                return LegacyXlsDrawingRecordKind.DrawingGroup;
            }

            if (type == (ushort)BiffRecordType.Drawing) {
                return LegacyXlsDrawingRecordKind.Drawing;
            }

            if (type == (ushort)BiffRecordType.Obj) {
                return LegacyXlsDrawingRecordKind.Object;
            }

            if (type == (ushort)BiffRecordType.Txo) {
                return LegacyXlsDrawingRecordKind.TextObject;
            }

            if (type == (ushort)BiffRecordType.ShapePropsStream) {
                return LegacyXlsDrawingRecordKind.ShapePropertiesStream;
            }

            if (type == (ushort)BiffRecordType.TextPropsStream) {
                return LegacyXlsDrawingRecordKind.TextPropertiesStream;
            }

            if (type == (ushort)BiffRecordType.RichTextStream) {
                return LegacyXlsDrawingRecordKind.RichTextStream;
            }

            if (type == (ushort)BiffRecordType.HfPicture) {
                return LegacyXlsDrawingRecordKind.HeaderFooterPicture;
            }

            return LegacyXlsDrawingRecordKind.PreserveOnly;
        }
    }
}
