using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffDrawingMetadataReader {
        internal static bool TryRead(
            BiffRecord record,
            string? sheetName,
            List<LegacyXlsDrawingRecord> records) {
            if (!BiffUnsupportedRecordDiagnostics.IsDrawingRecord(record.Type)) {
                return false;
            }

            TryReadObjectCommonData(record, out ushort? objectType, out ushort? objectId, out ushort? objectFlags);
            TryReadEscherHeader(record, out ushort? escherRecordType, out ushort? escherRecordInstance, out byte? escherRecordVersion, out uint? escherPayloadLength);
            ReadOfficeArtMetadata(
                record,
                out IReadOnlyList<LegacyXlsDrawingBlipStoreEntry> blipStoreEntries,
                out IReadOnlyList<LegacyXlsDrawingShape> shapeEntries,
                out IReadOnlyList<LegacyXlsDrawingAnchor> anchorEntries,
                out IReadOnlyList<LegacyXlsDrawingChildAnchor> childAnchorEntries,
                out IReadOnlyList<LegacyXlsDrawingOfficeArtRecord> officeArtRecords,
                out IReadOnlyList<LegacyXlsDrawingShapeProperty> shapeProperties);
            records.Add(new LegacyXlsDrawingRecord(
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
                shapeProperties: shapeProperties));
            return true;
        }

        private static void ReadOfficeArtMetadata(
            BiffRecord record,
            out IReadOnlyList<LegacyXlsDrawingBlipStoreEntry> blipStoreEntries,
            out IReadOnlyList<LegacyXlsDrawingShape> shapeEntries,
            out IReadOnlyList<LegacyXlsDrawingAnchor> anchorEntries,
            out IReadOnlyList<LegacyXlsDrawingChildAnchor> childAnchorEntries,
            out IReadOnlyList<LegacyXlsDrawingOfficeArtRecord> officeArtRecords,
            out IReadOnlyList<LegacyXlsDrawingShapeProperty> shapeProperties) {
            if (record.Type != (ushort)BiffRecordType.DrawingGroup && record.Type != (ushort)BiffRecordType.Drawing) {
                blipStoreEntries = Array.Empty<LegacyXlsDrawingBlipStoreEntry>();
                shapeEntries = Array.Empty<LegacyXlsDrawingShape>();
                anchorEntries = Array.Empty<LegacyXlsDrawingAnchor>();
                childAnchorEntries = Array.Empty<LegacyXlsDrawingChildAnchor>();
                officeArtRecords = Array.Empty<LegacyXlsDrawingOfficeArtRecord>();
                shapeProperties = Array.Empty<LegacyXlsDrawingShapeProperty>();
                return;
            }

            var blips = new List<LegacyXlsDrawingBlipStoreEntry>();
            var shapes = new List<LegacyXlsDrawingShape>();
            var anchors = new List<LegacyXlsDrawingAnchor>();
            var childAnchors = new List<LegacyXlsDrawingChildAnchor>();
            var records = new List<LegacyXlsDrawingOfficeArtRecord>();
            var properties = new List<LegacyXlsDrawingShapeProperty>();
            TryReadOfficeArtRecords(record.Payload, 0, record.Payload.Length, records, blips, shapes, anchors, childAnchors, properties, depth: 0);
            blipStoreEntries = blips;
            shapeEntries = shapes;
            anchorEntries = anchors;
            childAnchorEntries = childAnchors;
            officeArtRecords = records;
            shapeProperties = properties;
        }

        private static void TryReadOfficeArtRecords(
            byte[] payload,
            int startOffset,
            int endOffset,
            List<LegacyXlsDrawingOfficeArtRecord> officeArtRecords,
            List<LegacyXlsDrawingBlipStoreEntry> blipStoreEntries,
            List<LegacyXlsDrawingShape> shapeEntries,
            List<LegacyXlsDrawingAnchor> anchorEntries,
            List<LegacyXlsDrawingChildAnchor> childAnchorEntries,
            List<LegacyXlsDrawingShapeProperty> shapeProperties,
            int depth) {
            if (depth > 8) {
                return;
            }

            int offset = startOffset;
            while (offset + 8 <= endOffset) {
                ushort options = BiffRecordReader.ReadUInt16(payload, offset);
                ushort recordType = BiffRecordReader.ReadUInt16(payload, offset + 2);
                uint recordLength = BiffRecordReader.ReadUInt32(payload, offset + 4);
                if (recordLength > int.MaxValue || offset + 8 + (int)recordLength > endOffset) {
                    return;
                }

                int contentStart = offset + 8;
                int contentEnd = contentStart + (int)recordLength;
                byte version = checked((byte)(options & 0x000f));
                ushort instance = checked((ushort)(options >> 4));
                officeArtRecords.Add(new LegacyXlsDrawingOfficeArtRecord(recordType, instance, version, recordLength, depth));

                if (recordType == 0xF007 && TryReadBlipStoreEntry(payload, contentStart, contentEnd, instance, out LegacyXlsDrawingBlipStoreEntry? blipEntry)) {
                    blipStoreEntries.Add(blipEntry!);
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
                    TryReadOfficeArtRecords(payload, contentStart, contentEnd, officeArtRecords, blipStoreEntries, shapeEntries, anchorEntries, childAnchorEntries, shapeProperties, depth + 1);
                }

                offset = contentEnd;
            }
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
            uint sizeBytes = BiffRecordReader.ReadUInt32(payload, contentStart + 20);
            uint referenceCount = BiffRecordReader.ReadUInt32(payload, contentStart + 24);
            byte nameByteCount = payload[contentStart + 33];
            int embeddedOffset = contentStart + 36 + nameByteCount;
            ushort? embeddedRecordType = null;
            uint? embeddedPayloadLength = null;
            if (embeddedOffset + 8 <= contentEnd) {
                embeddedRecordType = BiffRecordReader.ReadUInt16(payload, embeddedOffset + 2);
                embeddedPayloadLength = BiffRecordReader.ReadUInt32(payload, embeddedOffset + 4);
            }

            entry = new LegacyXlsDrawingBlipStoreEntry(
                recordInstance,
                win32BlipType,
                macOsBlipType,
                sizeBytes,
                referenceCount,
                embeddedRecordType,
                embeddedPayloadLength);
            return true;
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
                if (isComplex) {
                    availableComplexDataLength = GetAvailableComplexDataLength(value, complexOffset, contentEnd);
                    complexOffset += availableComplexDataLength.Value;
                }

                properties.Add(new LegacyXlsDrawingShapeProperty(index, rawOperationId, value, availableComplexDataLength));
            }
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
            if (record.Type != (ushort)BiffRecordType.DrawingGroup &&
                record.Type != (ushort)BiffRecordType.Drawing) {
                return false;
            }

            if (payload.Length < 8) {
                return false;
            }

            ushort options = BiffRecordReader.ReadUInt16(payload, 0);
            recordVersion = checked((byte)(options & 0x000f));
            recordInstance = checked((ushort)(options >> 4));
            recordType = BiffRecordReader.ReadUInt16(payload, 2);
            payloadLength = BiffRecordReader.ReadUInt32(payload, 4);
            return true;
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

            return LegacyXlsDrawingRecordKind.PreserveOnly;
        }
    }
}
