using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffPivotTableMetadataReader {
        internal static bool TryRead(
            BiffRecord record,
            string? sheetName,
            List<LegacyXlsPivotTableRecord> records,
            List<LegacyXlsImportDiagnostic> diagnostics,
            BiffPivotTableMetadataReaderState? state = null) {
            if (!BiffUnsupportedRecordDiagnostics.IsPivotTableRecord(record.Type)) {
                return false;
            }

            LegacyXlsPivotTableRecord pivotRecord = CreateRecord(record, sheetName);
            records.Add(pivotRecord);
            try {
                state?.TryAttachGroupingRangeValue(record);
                switch (record.Type) {
                    case 0x00C1:
                        ReadDataItem(record, pivotRecord);
                        break;
                    case 0x00D7:
                        ReadGroupingRange(record, pivotRecord);
                        state?.TrackGroupingRange(record, pivotRecord);
                        break;
                    case 0x00FF:
                        ReadExtendedPivotField(record, pivotRecord);
                        break;
                    case 0x0864:
                        ReadAdditionalMetadata(record, pivotRecord);
                        break;
                }
            } catch (InvalidDataException) {
                // PivotTable import is currently preserve-only. Keep the typed record node,
                // but avoid turning partial metadata decode into an import warning.
            }

            return true;
        }

        private static LegacyXlsPivotTableRecord CreateRecord(BiffRecord record, string? sheetName) {
            LegacyXlsPivotTableRecordKind kind;
            switch (record.Type) {
                case 0x00B0:
                    kind = LegacyXlsPivotTableRecordKind.View;
                    break;
                case 0x00B1:
                    kind = LegacyXlsPivotTableRecordKind.Field;
                    break;
                case 0x00B2:
                    kind = LegacyXlsPivotTableRecordKind.Item;
                    break;
                case 0x00B4:
                    kind = LegacyXlsPivotTableRecordKind.FieldIndexList;
                    break;
                case 0x00B5:
                    kind = LegacyXlsPivotTableRecordKind.LineItem;
                    break;
                case 0x00B6:
                    kind = LegacyXlsPivotTableRecordKind.PageItem;
                    break;
                case 0x00C1:
                    kind = LegacyXlsPivotTableRecordKind.DataItem;
                    break;
                case 0x00C5:
                case 0x00C6:
                case 0x00C7:
                    kind = LegacyXlsPivotTableRecordKind.Cache;
                    break;
                case 0x00C8:
                case 0x00C9:
                case 0x00CA:
                case 0x00CB:
                case 0x00CC:
                case 0x00CD:
                case 0x00CE:
                    kind = LegacyXlsPivotTableRecordKind.CacheItem;
                    break;
                case 0x00CF:
                case 0x00D0:
                case 0x00D1:
                    kind = LegacyXlsPivotTableRecordKind.Table;
                    break;
                case 0x00D5:
                    kind = LegacyXlsPivotTableRecordKind.CacheStream;
                    break;
                case 0x00D7:
                    kind = LegacyXlsPivotTableRecordKind.GroupingRange;
                    break;
                case 0x00D8:
                    kind = LegacyXlsPivotTableRecordKind.Formula;
                    break;
                case 0x00EF:
                    kind = LegacyXlsPivotTableRecordKind.Rule;
                    break;
                case 0x00F0:
                    kind = LegacyXlsPivotTableRecordKind.CacheExtension;
                    break;
                case 0x00F1:
                    kind = LegacyXlsPivotTableRecordKind.Filter;
                    break;
                case 0x00F2:
                case 0x00F9:
                    kind = LegacyXlsPivotTableRecordKind.Format;
                    break;
                case 0x00F4:
                    kind = LegacyXlsPivotTableRecordKind.Item;
                    break;
                case 0x00F5:
                    kind = LegacyXlsPivotTableRecordKind.Field;
                    break;
                case 0x00F6:
                    kind = LegacyXlsPivotTableRecordKind.Selection;
                    break;
                case 0x00F7:
                    kind = LegacyXlsPivotTableRecordKind.PageItem;
                    break;
                case 0x00F8:
                case 0x0100:
                    kind = LegacyXlsPivotTableRecordKind.Formula;
                    break;
                case 0x00FF:
                    kind = LegacyXlsPivotTableRecordKind.ExtendedPivotField;
                    break;
                case 0x0122:
                    kind = LegacyXlsPivotTableRecordKind.CacheExtension;
                    break;
                case 0x0801:
                    kind = LegacyXlsPivotTableRecordKind.QueryTableTag;
                    break;
                case 0x0857:
                    kind = LegacyXlsPivotTableRecordKind.ViewLink;
                    break;
                case 0x0858:
                    kind = LegacyXlsPivotTableRecordKind.PivotChart;
                    break;
                case 0x0864:
                    kind = LegacyXlsPivotTableRecordKind.Additional;
                    break;
                default:
                    kind = LegacyXlsPivotTableRecordKind.PreserveOnly;
                    break;
            }

            return new LegacyXlsPivotTableRecord(
                kind,
                BiffUnsupportedRecordDiagnostics.GetBiffRecordName(record.Type),
                sheetName,
                record.Offset,
                record.Type,
                record.Payload.Length);
        }

        private static void ReadDataItem(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length < 14) {
                throw new InvalidDataException("The SXDI payload is shorter than the fixed data item header.");
            }

            short dataItemFieldIndex = BiffRecordReader.ReadInt16(payload, 0);
            short aggregationFunction = BiffRecordReader.ReadInt16(payload, 2);
            short displayCalculation = BiffRecordReader.ReadInt16(payload, 4);
            short displayCalculationFieldIndex = BiffRecordReader.ReadInt16(payload, 6);
            short displayCalculationItemIndex = BiffRecordReader.ReadInt16(payload, 8);
            ushort numberFormatId = BiffRecordReader.ReadUInt16(payload, 10);
            ushort nameLength = BiffRecordReader.ReadUInt16(payload, 12);
            string? name = null;
            if (nameLength != 0xFFFF && nameLength != 0) {
                int offset = 14;
                name = BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, nameLength);
            }

            pivotRecord.SetDataItem(
                dataItemFieldIndex,
                aggregationFunction,
                displayCalculation,
                displayCalculationFieldIndex,
                displayCalculationItemIndex,
                numberFormatId,
                name);
        }

        private static void ReadGroupingRange(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length < 2) {
                throw new InvalidDataException("The SXRng payload is shorter than the grouping flags.");
            }

            ushort flags = BiffRecordReader.ReadUInt16(payload, 0);
            bool autoStart = (flags & 0x0001) != 0;
            bool autoEnd = (flags & 0x0002) != 0;
            int groupingValue = (flags >> 2) & 0x0007;
            pivotRecord.SetGroupingRange(autoStart, autoEnd, (LegacyXlsPivotGroupingKind)groupingValue);
        }

        private static void ReadExtendedPivotField(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length < 2) {
                throw new InvalidDataException("The SXVDEx payload is shorter than the extended pivot field flags.");
            }

            ushort flags = BiffRecordReader.ReadUInt16(payload, 0);
            pivotRecord.SetExtendedPivotField(
                showAllItems: (flags & 0x0001) != 0,
                canDragToRow: (flags & 0x0002) != 0,
                canDragToColumn: (flags & 0x0004) != 0,
                canDragToPage: (flags & 0x0008) != 0,
                canDragToHide: (flags & 0x0010) != 0,
                preventDragToData: (flags & 0x0020) != 0,
                serverBased: (flags & 0x0080) != 0);
        }

        private static void ReadAdditionalMetadata(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length < 6) {
                throw new InvalidDataException("The SXAddl payload is shorter than the future-record header.");
            }

            ushort futureRecordType = BiffRecordReader.ReadUInt16(payload, 0);
            ushort futureFlags = BiffRecordReader.ReadUInt16(payload, 2);
            byte additionalClass = payload[4];
            byte additionalType = payload[5];
            uint? cacheId = null;
            if (additionalClass == 0x03 && additionalType == 0x00 && payload.Length >= 10) {
                cacheId = BiffRecordReader.ReadUInt32(payload, 6);
            }

            pivotRecord.SetAdditionalMetadata(
                futureRecordType,
                futureFlags,
                additionalClass,
                additionalType,
                cacheId);
        }
    }

    internal sealed class BiffPivotTableMetadataReaderState {
        private LegacyXlsPivotTableRecord? _pendingGroupingRange;
        private int _expectedOffset;
        private int _valueIndex;

        internal void TrackGroupingRange(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            _pendingGroupingRange = pivotRecord;
            _expectedOffset = record.Offset + 4 + record.Payload.Length;
            _valueIndex = 0;
        }

        internal void TryAttachGroupingRangeValue(BiffRecord record) {
            if (_pendingGroupingRange == null || record.Offset != _expectedOffset) {
                Reset();
                return;
            }

            try {
                if (_pendingGroupingRange.GroupingKind == LegacyXlsPivotGroupingKind.Numeric) {
                    TryAttachNumericGroupingValue(record);
                } else {
                    TryAttachDateGroupingValue(record);
                }
            } catch (InvalidDataException) {
                Reset();
            }
        }

        private void TryAttachNumericGroupingValue(BiffRecord record) {
            if (record.Type != 0x00C8 || _valueIndex >= 3) {
                Reset();
                return;
            }

            _pendingGroupingRange!.SetGroupingNumericValue(_valueIndex, BiffRecordReader.ReadDouble(record.Payload, 0));
            Advance(record);
        }

        private void TryAttachDateGroupingValue(BiffRecord record) {
            if (_valueIndex < 2) {
                if (record.Type != 0x00CD) {
                    Reset();
                    return;
                }

                _pendingGroupingRange!.SetGroupingDateValue(_valueIndex, ReadDateTimeValue(record.Payload));
                Advance(record);
                return;
            }

            if (_valueIndex == 2 && record.Type == 0x00CB) {
                _pendingGroupingRange!.SetGroupingDateInterval(BiffRecordReader.ReadInt16(record.Payload, 0));
                Advance(record);
                return;
            }

            Reset();
        }

        private void Advance(BiffRecord record) {
            _valueIndex++;
            if (_valueIndex >= 3) {
                Reset();
                return;
            }

            _expectedOffset = record.Offset + 4 + record.Payload.Length;
        }

        private void Reset() {
            _pendingGroupingRange = null;
            _expectedOffset = 0;
            _valueIndex = 0;
        }

        private static LegacyXlsPivotDateTimeValue ReadDateTimeValue(byte[] payload) {
            if (payload.Length < 8) {
                throw new InvalidDataException("The SXDtr payload is shorter than the date/time value.");
            }

            return new LegacyXlsPivotDateTimeValue(
                BiffRecordReader.ReadUInt16(payload, 0),
                BiffRecordReader.ReadUInt16(payload, 2),
                payload[4],
                payload[5],
                payload[6],
                payload[7]);
        }
    }
}
