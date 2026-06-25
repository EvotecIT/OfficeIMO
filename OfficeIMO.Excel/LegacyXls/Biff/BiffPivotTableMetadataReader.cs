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
                    case 0x00B0:
                        ReadView(record, pivotRecord);
                        break;
                    case 0x00B1:
                        ReadField(record, pivotRecord);
                        break;
                    case 0x00B2:
                        ReadItem(record, pivotRecord);
                        break;
                    case 0x00C1:
                        ReadDataItem(record, pivotRecord);
                        break;
                    case 0x00C5:
                        ReadCacheProperties(record, pivotRecord);
                        break;
                    case 0x00C8:
                        ReadCacheItemNumber(record, pivotRecord);
                        break;
                    case 0x00C9:
                        ReadCacheItemBoolean(record, pivotRecord);
                        break;
                    case 0x00CA:
                        ReadCacheItemError(record, pivotRecord);
                        break;
                    case 0x00CB:
                        ReadCacheItemInteger(record, pivotRecord);
                        break;
                    case 0x00CC:
                        ReadCacheItemString(record, pivotRecord);
                        break;
                    case 0x00CD:
                        ReadCacheItemDateTime(record, pivotRecord);
                        break;
                    case 0x00CE:
                        pivotRecord.SetCacheItemEmpty();
                        break;
                    case 0x00D5:
                        ReadCacheStreamId(record, pivotRecord);
                        break;
                    case 0x00D7:
                        ReadGroupingRange(record, pivotRecord);
                        state?.TrackGroupingRange(record, pivotRecord);
                        break;
                    case 0x0100:
                        ReadCalculatedItemFormula(record, pivotRecord);
                        break;
                    case 0x00E3:
                        ReadCacheSourceType(record, pivotRecord);
                        break;
                    case 0x00FF:
                        ReadExtendedPivotField(record, pivotRecord);
                        break;
                    case 0x0802:
                        ReadQueryTableTag(record, pivotRecord);
                        break;
                    case 0x0864:
                        ReadAdditionalMetadata(record, pivotRecord);
                        state?.TrackAdditionalRecord(pivotRecord);
                        break;
                }
            } catch (InvalidDataException) {
                // PivotTable import is currently preserve-only. Keep the typed record node,
                // but avoid turning partial metadata decode into an import warning.
            }

            return true;
        }

        private static void ReadView(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length < 44) {
                throw new InvalidDataException("The SxView payload is shorter than the fixed PivotTable view header.");
            }

            ushort firstRow = BiffRecordReader.ReadUInt16(payload, 0);
            ushort lastRow = BiffRecordReader.ReadUInt16(payload, 2);
            ushort firstColumn = BiffRecordReader.ReadUInt16(payload, 4);
            ushort lastColumn = BiffRecordReader.ReadUInt16(payload, 6);
            ushort firstHeaderRow = BiffRecordReader.ReadUInt16(payload, 8);
            ushort firstDataRow = BiffRecordReader.ReadUInt16(payload, 10);
            ushort firstDataColumn = BiffRecordReader.ReadUInt16(payload, 12);
            short cacheIndex = BiffRecordReader.ReadInt16(payload, 14);
            ushort dataAxis = BiffRecordReader.ReadUInt16(payload, 18);
            short dataPosition = BiffRecordReader.ReadInt16(payload, 20);
            short fieldCount = BiffRecordReader.ReadInt16(payload, 22);
            ushort rowFieldCount = BiffRecordReader.ReadUInt16(payload, 24);
            ushort columnFieldCount = BiffRecordReader.ReadUInt16(payload, 26);
            ushort pageFieldCount = BiffRecordReader.ReadUInt16(payload, 28);
            short dataFieldCount = BiffRecordReader.ReadInt16(payload, 30);
            ushort rowLineCount = BiffRecordReader.ReadUInt16(payload, 32);
            ushort columnLineCount = BiffRecordReader.ReadUInt16(payload, 34);
            ushort flags = BiffRecordReader.ReadUInt16(payload, 36);
            ushort autoFormatId = BiffRecordReader.ReadUInt16(payload, 38);
            ushort tableNameLength = BiffRecordReader.ReadUInt16(payload, 40);
            ushort dataNameLength = BiffRecordReader.ReadUInt16(payload, 42);
            int offset = 44;
            string tableName = tableNameLength == 0
                ? string.Empty
                : BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, tableNameLength);
            string dataName = dataNameLength == 0
                ? string.Empty
                : BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, dataNameLength);

            pivotRecord.SetView(
                firstRow,
                lastRow,
                firstColumn,
                lastColumn,
                firstHeaderRow,
                firstDataRow,
                firstDataColumn,
                cacheIndex,
                dataAxis,
                dataPosition,
                fieldCount,
                rowFieldCount,
                columnFieldCount,
                pageFieldCount,
                dataFieldCount,
                rowLineCount,
                columnLineCount,
                flags,
                autoFormatId,
                tableName,
                dataName);
        }

        private static void ReadField(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length < 10) {
                throw new InvalidDataException("The Sxvd payload is shorter than the fixed PivotTable field header.");
            }

            ushort axis = BiffRecordReader.ReadUInt16(payload, 0);
            ushort subtotalCount = BiffRecordReader.ReadUInt16(payload, 2);
            ushort subtotalFlags = BiffRecordReader.ReadUInt16(payload, 4);
            short itemCount = BiffRecordReader.ReadInt16(payload, 6);
            ushort nameLength = BiffRecordReader.ReadUInt16(payload, 8);
            string? name = null;
            if (nameLength != 0xFFFF && nameLength != 0) {
                int offset = 10;
                name = BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, nameLength);
            }

            pivotRecord.SetField(axis, subtotalFlags, subtotalCount, itemCount, name);
        }

        private static void ReadItem(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length < 8) {
                throw new InvalidDataException("The SXVI payload is shorter than the fixed PivotTable item header.");
            }

            short itemType = BiffRecordReader.ReadInt16(payload, 0);
            ushort flags = BiffRecordReader.ReadUInt16(payload, 2);
            short cacheIndex = BiffRecordReader.ReadInt16(payload, 4);
            ushort nameLength = BiffRecordReader.ReadUInt16(payload, 6);
            string? name = null;
            if (nameLength != 0xFFFF && nameLength != 0) {
                int offset = 8;
                name = BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, nameLength);
            }

            pivotRecord.SetItem(itemType, flags, cacheIndex, name);
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
                case 0x00E3:
                    kind = LegacyXlsPivotTableRecordKind.CacheSource;
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
                case 0x0802:
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

        private static void ReadCacheProperties(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length < 20) {
                throw new InvalidDataException("The SXDB payload is shorter than the fixed PivotCache header.");
            }

            int recordCount = BiffRecordReader.ReadInt32(payload, 0);
            ushort streamId = BiffRecordReader.ReadUInt16(payload, 4);
            ushort flags = BiffRecordReader.ReadUInt16(payload, 6);
            short sourceFieldCount = BiffRecordReader.ReadInt16(payload, 10);
            short totalFieldCount = BiffRecordReader.ReadInt16(payload, 12);
            ushort usedRecordCount = BiffRecordReader.ReadUInt16(payload, 14);
            ushort sourceType = BiffRecordReader.ReadUInt16(payload, 16);
            ushort refreshedByLength = BiffRecordReader.ReadUInt16(payload, 18);
            string? refreshedBy = null;
            if (refreshedByLength != 0xFFFF && refreshedByLength != 0) {
                int offset = 20;
                refreshedBy = BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, refreshedByLength);
            }

            pivotRecord.SetCacheProperties(
                recordCount,
                streamId,
                hasRecords: (flags & 0x0001) != 0,
                invalid: (flags & 0x0002) != 0,
                refreshOnLoad: (flags & 0x0004) != 0,
                optimizeMemory: (flags & 0x0008) != 0,
                backgroundQuery: (flags & 0x0010) != 0,
                enableRefresh: (flags & 0x0020) != 0,
                sourceFieldCount,
                totalFieldCount,
                usedRecordCount,
                sourceType,
                refreshedBy);
        }

        private static void ReadQueryTableTag(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length < 19) {
                throw new InvalidDataException("The QsiSXTag payload is shorter than the fixed tag header.");
            }

            ushort futureRecordType = BiffRecordReader.ReadUInt16(payload, 0);
            ushort futureFlags = BiffRecordReader.ReadUInt16(payload, 2);
            bool relatesToPivotTable = BiffRecordReader.ReadUInt16(payload, 4) != 0;
            ushort flags = BiffRecordReader.ReadUInt16(payload, 6);
            uint futureOptions = BiffRecordReader.ReadUInt32(payload, 8);
            byte lastUpdatedVersion = payload[12];
            byte updatableMinimumVersion = payload[13];
            byte nameOffsetMarker = payload[14];
            int offset = 16;
            string name = BiffStringReader.ReadUnicodeString(payload, ref offset);
            if (offset + 2 > payload.Length) {
                throw new InvalidDataException("The QsiSXTag payload is missing the trailing unused field.");
            }

            pivotRecord.SetQueryTableTag(
                futureRecordType,
                futureFlags,
                relatesToPivotTable,
                refreshEnabled: (flags & 0x0001) != 0,
                cacheInvalid: (flags & 0x0002) != 0,
                tensorEx: (flags & 0x0004) != 0,
                futureOptions,
                lastUpdatedVersion,
                updatableMinimumVersion,
                nameOffsetMarker,
                name,
                BiffRecordReader.ReadUInt16(payload, offset));
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

        private static void ReadCacheItemNumber(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            if (record.Payload.Length < 8) {
                throw new InvalidDataException("The SXNum payload is shorter than the numeric value.");
            }

            pivotRecord.SetCacheItemNumber(BiffRecordReader.ReadDouble(record.Payload, 0));
        }

        private static void ReadCacheItemBoolean(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            if (record.Payload.Length < 2) {
                throw new InvalidDataException("The SxBool payload is shorter than the Boolean value.");
            }

            pivotRecord.SetCacheItemBoolean(BiffRecordReader.ReadUInt16(record.Payload, 0) != 0);
        }

        private static void ReadCacheItemError(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            if (record.Payload.Length < 2) {
                throw new InvalidDataException("The SxErr payload is shorter than the error value.");
            }

            ushort errorCode = BiffRecordReader.ReadUInt16(record.Payload, 0);
            string errorText = errorCode <= byte.MaxValue
                ? BiffErrorValue.ToText((byte)errorCode)
                : $"#ERR({errorCode})";
            pivotRecord.SetCacheItemError(errorCode, errorText);
        }

        private static void ReadCacheItemInteger(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            if (record.Payload.Length < 2) {
                throw new InvalidDataException("The SXInt payload is shorter than the integer value.");
            }

            pivotRecord.SetCacheItemInteger(BiffRecordReader.ReadInt16(record.Payload, 0));
        }

        private static void ReadCacheItemString(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length < 2) {
                throw new InvalidDataException("The SXString payload is shorter than the character count.");
            }

            ushort characterCount = BiffRecordReader.ReadUInt16(payload, 0);
            string? value = null;
            if (characterCount != 0xFFFF) {
                int offset = 2;
                value = BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, characterCount);
            }

            pivotRecord.SetCacheItemString(value);
        }

        private static void ReadCacheItemDateTime(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            pivotRecord.SetCacheItemDateTime(ReadDateTimeValue(record.Payload));
        }

        private static void ReadCacheStreamId(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            if (record.Payload.Length < 2) {
                throw new InvalidDataException("The SXStreamID payload is shorter than the stream identifier.");
            }

            pivotRecord.SetCacheStream(BiffRecordReader.ReadUInt16(record.Payload, 0));
        }

        private static void ReadCacheSourceType(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            if (record.Payload.Length < 2) {
                throw new InvalidDataException("The SXVS payload is shorter than the cache source type.");
            }

            pivotRecord.SetCacheSourceType(BiffRecordReader.ReadUInt16(record.Payload, 0));
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

        private static void ReadCalculatedItemFormula(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length != 4) {
                throw new InvalidDataException("The SXFormula payload length does not match the cache-field index structure.");
            }

            pivotRecord.SetCalculatedItemFormula(
                BiffRecordReader.ReadUInt16(payload, 0),
                BiffRecordReader.ReadInt16(payload, 2));
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

    internal sealed class BiffPivotTableMetadataReaderState {
        private LegacyXlsPivotTableRecord? _pendingGroupingRange;
        private int _expectedOffset;
        private int _valueIndex;
        private int _additionalDepth;
        private int _additionalSequence;

        internal void TrackAdditionalRecord(LegacyXlsPivotTableRecord pivotRecord) {
            if (!pivotRecord.AdditionalClass.HasValue || !pivotRecord.AdditionalType.HasValue) {
                return;
            }

            _additionalSequence++;
            int depthBefore = _additionalDepth;
            string transition;

            if (pivotRecord.AdditionalType == 0x00) {
                _additionalDepth++;
                transition = "BeginClass";
            } else if (pivotRecord.AdditionalType == 0xFF) {
                if (_additionalDepth == 0) {
                    transition = "UnmatchedEndClass";
                } else {
                    _additionalDepth--;
                    transition = "EndClass";
                }
            } else {
                transition = _additionalDepth == 0 ? "OutsideClass" : "InsideClass";
            }

            pivotRecord.SetAdditionalClassNesting(_additionalSequence, depthBefore, _additionalDepth, transition);
        }

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
