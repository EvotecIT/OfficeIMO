namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a preserve-only PivotTable BIFF record and any shallow fields decoded from it.
    /// </summary>
    public sealed class LegacyXlsPivotTableRecord {
        /// <summary>
        /// Creates PivotTable BIFF record metadata.
        /// </summary>
        public LegacyXlsPivotTableRecord(
            LegacyXlsPivotTableRecordKind kind,
            string recordName,
            string? sheetName,
            int recordOffset,
            ushort recordType,
            int payloadLength) {
            if (payloadLength < 0) {
                throw new ArgumentOutOfRangeException(nameof(payloadLength));
            }

            Kind = kind;
            RecordName = recordName ?? throw new ArgumentNullException(nameof(recordName));
            SheetName = sheetName;
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
        }

        /// <summary>Gets the decoded PivotTable metadata kind.</summary>
        public LegacyXlsPivotTableRecordKind Kind { get; }

        /// <summary>Gets the BIFF record name.</summary>
        public string RecordName { get; }

        /// <summary>Gets the worksheet or sheet entry name associated with the record, when known.</summary>
        public string? SheetName { get; }

        /// <summary>Gets the byte offset of the BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type identifier.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the BIFF record payload length in bytes.</summary>
        public int PayloadLength { get; }

        /// <summary>Gets the pivot field index for an SXDI data item, when decoded.</summary>
        public short? DataItemFieldIndex { get; private set; }

        /// <summary>Gets the aggregation function identifier for an SXDI data item, when decoded.</summary>
        public short? AggregationFunction { get; private set; }

        /// <summary>Gets the decoded aggregation function for an SXDI data item, when the identifier is known.</summary>
        public LegacyXlsPivotAggregationFunction? AggregationFunctionKind { get; private set; }

        /// <summary>Gets the aggregation function name for an SXDI data item, or a stable raw identifier for unknown values.</summary>
        public string? AggregationFunctionName { get; private set; }

        /// <summary>Gets the display calculation identifier for an SXDI data item, when decoded.</summary>
        public short? DisplayCalculation { get; private set; }

        /// <summary>Gets the decoded display calculation for an SXDI data item, when the identifier is known.</summary>
        public LegacyXlsPivotDisplayCalculation? DisplayCalculationKind { get; private set; }

        /// <summary>Gets the display calculation name for an SXDI data item, or a stable raw identifier for unknown values.</summary>
        public string? DisplayCalculationName { get; private set; }

        /// <summary>Gets the pivot field index used by an SXDI display calculation, when decoded.</summary>
        public short? DisplayCalculationFieldIndex { get; private set; }

        /// <summary>Gets the pivot item index used by an SXDI display calculation, when decoded.</summary>
        public short? DisplayCalculationItemIndex { get; private set; }

        /// <summary>Gets the number format identifier stored in an SXDI data item, when decoded.</summary>
        public ushort? NumberFormatId { get; private set; }

        /// <summary>Gets the optional custom SXDI data item name, when decoded.</summary>
        public string? Name { get; private set; }

        /// <summary>Gets the PivotCache stream identifier, when decoded from SXStreamID or SXDB.</summary>
        public ushort? CacheStreamId { get; private set; }

        /// <summary>Gets the four-character PivotCache stream name implied by the stream identifier.</summary>
        public string? CacheStreamName { get; private set; }

        /// <summary>Gets the number of cache records declared by SXDB, when decoded.</summary>
        public int? CacheRecordCount { get; private set; }

        /// <summary>Gets whether SXDB declares cached records exist for this PivotCache.</summary>
        public bool? CacheHasRecords { get; private set; }

        /// <summary>Gets whether SXDB marks the PivotCache records as invalid.</summary>
        public bool? CacheInvalid { get; private set; }

        /// <summary>Gets whether SXDB requests refresh when the workbook is loaded.</summary>
        public bool? CacheRefreshOnLoad { get; private set; }

        /// <summary>Gets whether SXDB requests optimized cache storage.</summary>
        public bool? CacheOptimizeMemory { get; private set; }

        /// <summary>Gets whether SXDB declares background refresh for external cache data.</summary>
        public bool? CacheBackgroundQuery { get; private set; }

        /// <summary>Gets whether SXDB declares that cache refresh is enabled.</summary>
        public bool? CacheEnableRefresh { get; private set; }

        /// <summary>Gets the source-data field count declared by SXDB, when decoded.</summary>
        public short? CacheSourceFieldCount { get; private set; }

        /// <summary>Gets the total PivotCache field count declared by SXDB, when decoded.</summary>
        public short? CacheTotalFieldCount { get; private set; }

        /// <summary>Gets the cache record count used to calculate the PivotTable report, when decoded.</summary>
        public ushort? CacheUsedRecordCount { get; private set; }

        /// <summary>Gets the raw PivotCache source-data type declared by SXDB or SXVS.</summary>
        public ushort? CacheSourceType { get; private set; }

        /// <summary>Gets the decoded PivotCache source-data type, when known.</summary>
        public LegacyXlsPivotCacheSourceType? CacheSourceTypeKind { get; private set; }

        /// <summary>Gets the source-data type name, or a stable raw identifier for unknown values.</summary>
        public string? CacheSourceTypeName { get; private set; }

        /// <summary>Gets the optional user name stored in SXDB for the last cache refresh.</summary>
        public string? CacheRefreshedBy { get; private set; }

        /// <summary>Gets whether an SXRng record recalculates its starting value from source data.</summary>
        public bool? AutoStart { get; private set; }

        /// <summary>Gets whether an SXRng record recalculates its ending value from source data.</summary>
        public bool? AutoEnd { get; private set; }

        /// <summary>Gets the grouping criteria stored in an SXRng record, when decoded.</summary>
        public LegacyXlsPivotGroupingKind? GroupingKind { get; private set; }

        /// <summary>Gets the numeric grouping start value attached to an SXRng record, when decoded from the following SXNum records.</summary>
        public double? GroupingNumericStart { get; private set; }

        /// <summary>Gets the numeric grouping end value attached to an SXRng record, when decoded from the following SXNum records.</summary>
        public double? GroupingNumericEnd { get; private set; }

        /// <summary>Gets the numeric grouping interval attached to an SXRng record, when decoded from the following SXNum records.</summary>
        public double? GroupingNumericInterval { get; private set; }

        /// <summary>Gets the date grouping start value attached to an SXRng record, when decoded from the following SXDtr records.</summary>
        public LegacyXlsPivotDateTimeValue? GroupingDateStart { get; private set; }

        /// <summary>Gets the date grouping end value attached to an SXRng record, when decoded from the following SXDtr records.</summary>
        public LegacyXlsPivotDateTimeValue? GroupingDateEnd { get; private set; }

        /// <summary>Gets the date grouping interval attached to an SXRng record, when decoded from the following SXInt record.</summary>
        public short? GroupingDateInterval { get; private set; }

        /// <summary>Gets whether an SXVDEx record requests showing all items.</summary>
        public bool? ShowAllItems { get; private set; }

        /// <summary>Gets whether an SXVDEx record allows dragging the field to rows.</summary>
        public bool? CanDragToRow { get; private set; }

        /// <summary>Gets whether an SXVDEx record allows dragging the field to columns.</summary>
        public bool? CanDragToColumn { get; private set; }

        /// <summary>Gets whether an SXVDEx record allows dragging the field to pages.</summary>
        public bool? CanDragToPage { get; private set; }

        /// <summary>Gets whether an SXVDEx record allows hiding the field from the view.</summary>
        public bool? CanDragToHide { get; private set; }

        /// <summary>Gets whether an SXVDEx record prevents dragging the field to data values.</summary>
        public bool? PreventDragToData { get; private set; }

        /// <summary>Gets whether an SXVDEx record marks the pivot field as server-based.</summary>
        public bool? ServerBased { get; private set; }

        /// <summary>Gets the future-record type stored in an SXAddl header, when decoded.</summary>
        public ushort? AdditionalFutureRecordType { get; private set; }

        /// <summary>Gets the future-record flags stored in an SXAddl header, when decoded.</summary>
        public ushort? AdditionalFutureFlags { get; private set; }

        /// <summary>Gets the SXAddl class byte, when decoded.</summary>
        public byte? AdditionalClass { get; private set; }

        /// <summary>Gets the decoded SXAddl class name, or a stable raw identifier for unknown values.</summary>
        public string? AdditionalClassName { get; private set; }

        /// <summary>Gets the SXAddl detail type byte, when decoded.</summary>
        public byte? AdditionalType { get; private set; }

        /// <summary>Gets the decoded SXAddl detail type name, or a stable raw identifier for unknown values.</summary>
        public string? AdditionalTypeName { get; private set; }

        /// <summary>Gets the PivotCache identifier carried by an SXAddl SxcCache/SXDId record, when decoded.</summary>
        public uint? AdditionalCacheId { get; private set; }

        /// <summary>Gets the value kind for a PivotCache cache item record, when decoded.</summary>
        public LegacyXlsPivotCacheItemKind? CacheItemKind { get; private set; }

        /// <summary>Gets the value kind name for a PivotCache cache item record, when decoded.</summary>
        public string? CacheItemKindName { get; private set; }

        /// <summary>Gets the numeric value for an SXNum PivotCache item, when decoded.</summary>
        public double? CacheItemNumericValue { get; private set; }

        /// <summary>Gets the signed integer value for an SXInt PivotCache item, when decoded.</summary>
        public short? CacheItemIntegerValue { get; private set; }

        /// <summary>Gets the Boolean value for an SxBool PivotCache item, when decoded.</summary>
        public bool? CacheItemBooleanValue { get; private set; }

        /// <summary>Gets the error code for an SxErr PivotCache item, when decoded.</summary>
        public ushort? CacheItemErrorCode { get; private set; }

        /// <summary>Gets the error text for an SxErr PivotCache item, when decoded.</summary>
        public string? CacheItemErrorText { get; private set; }

        /// <summary>Gets the string value for an SXString PivotCache item, when decoded.</summary>
        public string? CacheItemStringValue { get; private set; }

        /// <summary>Gets the date/time value for an SXDtr PivotCache item, when decoded.</summary>
        public LegacyXlsPivotDateTimeValue? CacheItemDateTimeValue { get; private set; }

        /// <summary>Gets whether this record is an empty SxNil PivotCache item.</summary>
        public bool IsEmptyCacheItem => CacheItemKind == LegacyXlsPivotCacheItemKind.Empty;

        internal void SetDataItem(
            short dataItemFieldIndex,
            short aggregationFunction,
            short displayCalculation,
            short displayCalculationFieldIndex,
            short displayCalculationItemIndex,
            ushort numberFormatId,
            string? name) {
            DataItemFieldIndex = dataItemFieldIndex;
            AggregationFunction = aggregationFunction;
            AggregationFunctionKind = TryGetAggregationFunctionKind(aggregationFunction);
            AggregationFunctionName = AggregationFunctionKind?.ToString() ?? $"AggregationFunction:{aggregationFunction}";
            DisplayCalculation = displayCalculation;
            DisplayCalculationKind = TryGetDisplayCalculationKind(displayCalculation);
            DisplayCalculationName = DisplayCalculationKind?.ToString() ?? $"DisplayCalculation:{displayCalculation}";
            DisplayCalculationFieldIndex = displayCalculationFieldIndex;
            DisplayCalculationItemIndex = displayCalculationItemIndex;
            NumberFormatId = numberFormatId;
            Name = name;
        }

        internal void SetCacheStream(ushort streamId) {
            CacheStreamId = streamId;
            CacheStreamName = streamId.ToString("X4");
        }

        internal void SetCacheSourceType(ushort sourceType) {
            CacheSourceType = sourceType;
            CacheSourceTypeKind = TryGetCacheSourceTypeKind(sourceType);
            CacheSourceTypeName = CacheSourceTypeKind?.ToString() ?? $"SourceType:0x{sourceType:X4}";
        }

        internal void SetCacheProperties(
            int recordCount,
            ushort streamId,
            bool hasRecords,
            bool invalid,
            bool refreshOnLoad,
            bool optimizeMemory,
            bool backgroundQuery,
            bool enableRefresh,
            short sourceFieldCount,
            short totalFieldCount,
            ushort usedRecordCount,
            ushort sourceType,
            string? refreshedBy) {
            CacheRecordCount = recordCount;
            SetCacheStream(streamId);
            CacheHasRecords = hasRecords;
            CacheInvalid = invalid;
            CacheRefreshOnLoad = refreshOnLoad;
            CacheOptimizeMemory = optimizeMemory;
            CacheBackgroundQuery = backgroundQuery;
            CacheEnableRefresh = enableRefresh;
            CacheSourceFieldCount = sourceFieldCount;
            CacheTotalFieldCount = totalFieldCount;
            CacheUsedRecordCount = usedRecordCount;
            SetCacheSourceType(sourceType);
            CacheRefreshedBy = refreshedBy;
        }

        internal void SetGroupingRange(bool autoStart, bool autoEnd, LegacyXlsPivotGroupingKind groupingKind) {
            AutoStart = autoStart;
            AutoEnd = autoEnd;
            GroupingKind = groupingKind;
        }

        internal void SetGroupingNumericValue(int valueIndex, double value) {
            switch (valueIndex) {
                case 0:
                    GroupingNumericStart = value;
                    break;
                case 1:
                    GroupingNumericEnd = value;
                    break;
                case 2:
                    GroupingNumericInterval = value;
                    break;
            }
        }

        internal void SetGroupingDateValue(int valueIndex, LegacyXlsPivotDateTimeValue value) {
            if (valueIndex == 0) {
                GroupingDateStart = value;
            } else if (valueIndex == 1) {
                GroupingDateEnd = value;
            }
        }

        internal void SetGroupingDateInterval(short interval) {
            GroupingDateInterval = interval;
        }

        internal void SetExtendedPivotField(
            bool showAllItems,
            bool canDragToRow,
            bool canDragToColumn,
            bool canDragToPage,
            bool canDragToHide,
            bool preventDragToData,
            bool serverBased) {
            ShowAllItems = showAllItems;
            CanDragToRow = canDragToRow;
            CanDragToColumn = canDragToColumn;
            CanDragToPage = canDragToPage;
            CanDragToHide = canDragToHide;
            PreventDragToData = preventDragToData;
            ServerBased = serverBased;
        }

        internal void SetAdditionalMetadata(
            ushort futureRecordType,
            ushort futureFlags,
            byte additionalClass,
            byte additionalType,
            uint? cacheId) {
            AdditionalFutureRecordType = futureRecordType;
            AdditionalFutureFlags = futureFlags;
            AdditionalClass = additionalClass;
            AdditionalClassName = GetAdditionalClassName(additionalClass);
            AdditionalType = additionalType;
            AdditionalTypeName = GetAdditionalTypeName(additionalClass, additionalType);
            AdditionalCacheId = cacheId;
        }

        internal void SetCacheItemEmpty() {
            CacheItemKind = LegacyXlsPivotCacheItemKind.Empty;
            CacheItemKindName = LegacyXlsPivotCacheItemKind.Empty.ToString();
        }

        internal void SetCacheItemNumber(double value) {
            CacheItemKind = LegacyXlsPivotCacheItemKind.Number;
            CacheItemKindName = LegacyXlsPivotCacheItemKind.Number.ToString();
            CacheItemNumericValue = value;
        }

        internal void SetCacheItemBoolean(bool value) {
            CacheItemKind = LegacyXlsPivotCacheItemKind.Boolean;
            CacheItemKindName = LegacyXlsPivotCacheItemKind.Boolean.ToString();
            CacheItemBooleanValue = value;
        }

        internal void SetCacheItemError(ushort errorCode, string errorText) {
            CacheItemKind = LegacyXlsPivotCacheItemKind.Error;
            CacheItemKindName = LegacyXlsPivotCacheItemKind.Error.ToString();
            CacheItemErrorCode = errorCode;
            CacheItemErrorText = errorText;
        }

        internal void SetCacheItemInteger(short value) {
            CacheItemKind = LegacyXlsPivotCacheItemKind.Integer;
            CacheItemKindName = LegacyXlsPivotCacheItemKind.Integer.ToString();
            CacheItemIntegerValue = value;
        }

        internal void SetCacheItemString(string? value) {
            CacheItemKind = LegacyXlsPivotCacheItemKind.String;
            CacheItemKindName = LegacyXlsPivotCacheItemKind.String.ToString();
            CacheItemStringValue = value;
        }

        internal void SetCacheItemDateTime(LegacyXlsPivotDateTimeValue value) {
            CacheItemKind = LegacyXlsPivotCacheItemKind.DateTime;
            CacheItemKindName = LegacyXlsPivotCacheItemKind.DateTime.ToString();
            CacheItemDateTimeValue = value;
        }

        private static LegacyXlsPivotAggregationFunction? TryGetAggregationFunctionKind(short value) {
            return value >= 0 && value <= 10 ? (LegacyXlsPivotAggregationFunction)value : null;
        }

        private static LegacyXlsPivotDisplayCalculation? TryGetDisplayCalculationKind(short value) {
            return value >= 0 && value <= 8 ? (LegacyXlsPivotDisplayCalculation)value : null;
        }

        private static LegacyXlsPivotCacheSourceType? TryGetCacheSourceTypeKind(ushort value) {
            switch (value) {
                case 0x0001:
                    return LegacyXlsPivotCacheSourceType.Sheet;
                case 0x0002:
                    return LegacyXlsPivotCacheSourceType.External;
                case 0x0004:
                    return LegacyXlsPivotCacheSourceType.Consolidation;
                case 0x0010:
                    return LegacyXlsPivotCacheSourceType.Scenario;
                default:
                    return null;
            }
        }

        private static string GetAdditionalClassName(byte value) {
            switch (value) {
                case 0x00:
                    return "SxcView";
                case 0x01:
                    return "SxcField";
                case 0x02:
                    return "SxcHierarchy";
                case 0x03:
                    return "SxcCache";
                case 0x04:
                    return "SxcCacheField";
                case 0x05:
                    return "SxcQsi";
                case 0x06:
                    return "SxcQuery";
                case 0x07:
                    return "SxcGrpLevel";
                case 0x08:
                    return "SxcGroup";
                case 0x09:
                    return "SxcCacheItem";
                case 0x0C:
                    return "SxcSXRule";
                case 0x0D:
                    return "SxcSXFilt";
                case 0x10:
                    return "SxcSXDH";
                case 0x12:
                    return "SxcAutoSort";
                case 0x13:
                    return "SxcSXMgs";
                case 0x14:
                    return "SxcSXMg";
                case 0x17:
                    return "SxcField12";
                case 0x1A:
                    return "SxcSXCondFmts";
                case 0x1B:
                    return "SxcSXCondFmt";
                case 0x1C:
                    return "SxcSXFilters12";
                case 0x1D:
                    return "SxcSXFilter12";
                default:
                    return $"Sxc:0x{value:X2}";
            }
        }

        private static string GetAdditionalTypeName(byte additionalClass, byte value) {
            if (value == 0x00) {
                return "SXDId";
            }

            if (additionalClass == 0x00 && value == 0x02) {
                return "SXDVer10Info";
            }

            if ((additionalClass == 0x00 || additionalClass == 0x17) && value == 0x19) {
                return "SXDVer12Info";
            }

            if (value == 0xFF) {
                return "SXDEnd";
            }

            return $"SXD:0x{value:X2}";
        }
    }
}
