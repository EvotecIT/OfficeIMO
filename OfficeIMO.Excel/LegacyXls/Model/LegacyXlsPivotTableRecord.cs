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

        /// <summary>Gets the A1 range covered by an SxView PivotTable view, when decoded.</summary>
        public string? ViewRange { get; private set; }

        /// <summary>Gets the zero-based first row covered by an SxView PivotTable view, when decoded.</summary>
        public ushort? ViewFirstRow { get; private set; }

        /// <summary>Gets the zero-based last row covered by an SxView PivotTable view, when decoded.</summary>
        public ushort? ViewLastRow { get; private set; }

        /// <summary>Gets the zero-based first column covered by an SxView PivotTable view, when decoded.</summary>
        public ushort? ViewFirstColumn { get; private set; }

        /// <summary>Gets the zero-based last column covered by an SxView PivotTable view, when decoded.</summary>
        public ushort? ViewLastColumn { get; private set; }

        /// <summary>Gets the first row of the row area from an SxView PivotTable view, when decoded.</summary>
        public ushort? ViewFirstHeaderRow { get; private set; }

        /// <summary>Gets the first data row from an SxView PivotTable view, when decoded.</summary>
        public ushort? ViewFirstDataRow { get; private set; }

        /// <summary>Gets the first data column from an SxView PivotTable view, when decoded.</summary>
        public ushort? ViewFirstDataColumn { get; private set; }

        /// <summary>Gets the PivotCache index referenced by an SxView PivotTable view, when decoded.</summary>
        public short? ViewCacheIndex { get; private set; }

        /// <summary>Gets the default data-axis name decoded from sxaxis4Data in an SxView record.</summary>
        public string? ViewDataAxisName { get; private set; }

        /// <summary>Gets the data field position from an SxView PivotTable view, when decoded.</summary>
        public short? ViewDataPosition { get; private set; }

        /// <summary>Gets the declared pivot field count from an SxView record.</summary>
        public short? ViewFieldCount { get; private set; }

        /// <summary>Gets the row-axis field count from an SxView record.</summary>
        public ushort? ViewRowFieldCount { get; private set; }

        /// <summary>Gets the column-axis field count from an SxView record.</summary>
        public ushort? ViewColumnFieldCount { get; private set; }

        /// <summary>Gets the page-axis field count from an SxView record.</summary>
        public ushort? ViewPageFieldCount { get; private set; }

        /// <summary>Gets the data field count from an SxView record.</summary>
        public short? ViewDataFieldCount { get; private set; }

        /// <summary>Gets the row-area pivot line count from an SxView record.</summary>
        public ushort? ViewRowLineCount { get; private set; }

        /// <summary>Gets the column-area pivot line count from an SxView record.</summary>
        public ushort? ViewColumnLineCount { get; private set; }

        /// <summary>Gets whether an SxView record declares row grand totals.</summary>
        public bool? ViewRowGrandTotals { get; private set; }

        /// <summary>Gets whether an SxView record declares column grand totals.</summary>
        public bool? ViewColumnGrandTotals { get; private set; }

        /// <summary>Gets whether an SxView record declares AutoFormat is applied.</summary>
        public bool? ViewAutoFormat { get; private set; }

        /// <summary>Gets the AutoFormat identifier from an SxView record.</summary>
        public ushort? ViewAutoFormatId { get; private set; }

        /// <summary>Gets the PivotTable name carried by an SxView record.</summary>
        public string? ViewTableName { get; private set; }

        /// <summary>Gets the data field caption carried by an SxView record.</summary>
        public string? ViewDataName { get; private set; }

        /// <summary>Gets the decoded pivot field axis for an Sxvd record.</summary>
        public string? FieldAxisName { get; private set; }

        /// <summary>Gets the declared subtotal count for an Sxvd record.</summary>
        public ushort? FieldSubtotalCount { get; private set; }

        /// <summary>Gets the raw subtotal-function flags for an Sxvd record.</summary>
        public ushort? FieldSubtotalFlags { get; private set; }

        /// <summary>Gets decoded subtotal-function names for an Sxvd record.</summary>
        public IReadOnlyList<string> FieldSubtotalFunctionNames { get; private set; } = Array.Empty<string>();

        /// <summary>Gets the declared pivot item count for an Sxvd record.</summary>
        public short? FieldItemCount { get; private set; }

        /// <summary>Gets the optional pivot field caption carried by an Sxvd record.</summary>
        public string? FieldName { get; private set; }

        /// <summary>Gets the pivot field indexes referenced by an SxIvd row or column field-index list.</summary>
        public IReadOnlyList<short> FieldIndexReferences { get; private set; } = Array.Empty<short>();

        /// <summary>Gets the row or column line items decoded from an SXLI record.</summary>
        public IReadOnlyList<LegacyXlsPivotLineItem> LineItems { get; private set; } = Array.Empty<LegacyXlsPivotLineItem>();

        /// <summary>Gets the page-axis pivot item selectors decoded from an SXPI record.</summary>
        public IReadOnlyList<LegacyXlsPivotPageItem> PageItems { get; private set; } = Array.Empty<LegacyXlsPivotPageItem>();

        /// <summary>Gets the raw item type stored in an SXVI PivotTable item record.</summary>
        public short? ItemType { get; private set; }

        /// <summary>Gets the decoded item type stored in an SXVI PivotTable item record, when known.</summary>
        public LegacyXlsPivotItemType? ItemTypeKind { get; private set; }

        /// <summary>Gets the item type name for an SXVI PivotTable item record, or a stable raw identifier for unknown values.</summary>
        public string? ItemTypeName { get; private set; }

        /// <summary>Gets whether an SXVI PivotTable item is hidden.</summary>
        public bool? ItemHidden { get; private set; }

        /// <summary>Gets whether an SXVI PivotTable item hides detail.</summary>
        public bool? ItemHideDetail { get; private set; }

        /// <summary>Gets whether an SXVI PivotTable item represents a calculated item formula.</summary>
        public bool? ItemFormula { get; private set; }

        /// <summary>Gets whether an SXVI PivotTable item is missing from the cache.</summary>
        public bool? ItemMissing { get; private set; }

        /// <summary>Gets the PivotCache item index referenced by an SXVI PivotTable item.</summary>
        public short? ItemCacheIndex { get; private set; }

        /// <summary>Gets a stable cache-index name for an SXVI PivotTable item.</summary>
        public string? ItemCacheIndexName { get; private set; }

        /// <summary>Gets the optional caption carried by an SXVI PivotTable item.</summary>
        public string? ItemName { get; private set; }

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

        /// <summary>Gets the reserved SXFormula field, when decoded for a calculated-item formula scope.</summary>
        public ushort? CalculatedItemFormulaReserved { get; private set; }

        /// <summary>Gets the cache field index targeted by an SXFormula calculated-item formula scope.</summary>
        public short? CalculatedItemFormulaCacheFieldIndex { get; private set; }

        /// <summary>Gets whether SXFormula calculated-item formula scope metadata was decoded.</summary>
        public bool HasCalculatedItemFormulaScope => CalculatedItemFormulaCacheFieldIndex.HasValue;

        /// <summary>Gets whether SXFormula applies the calculated-item formula to all cache fields.</summary>
        public bool CalculatedItemFormulaAppliesToAllCacheFields => CalculatedItemFormulaCacheFieldIndex == -1;

        /// <summary>Gets a stable calculated-item formula scope name derived from SXFormula.</summary>
        public string? CalculatedItemFormulaScopeName { get; private set; }

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

        /// <summary>Gets the one-based SXAddl sequence index within the scanned PivotTable scope.</summary>
        public int? AdditionalSequenceIndex { get; private set; }

        /// <summary>Gets the SXAddl class nesting depth before this record is applied.</summary>
        public int? AdditionalClassDepthBefore { get; private set; }

        /// <summary>Gets the SXAddl class nesting depth after this record is applied.</summary>
        public int? AdditionalClassDepthAfter { get; private set; }

        /// <summary>Gets the shallow SXAddl class-stack transition represented by this record.</summary>
        public string? AdditionalClassTransition { get; private set; }

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

        /// <summary>Gets the future-record type stored in a QsiSXTag header, when decoded.</summary>
        public ushort? QueryTableTagFutureRecordType { get; private set; }

        /// <summary>Gets the future-record flags stored in a QsiSXTag header, when decoded.</summary>
        public ushort? QueryTableTagFutureFlags { get; private set; }

        /// <summary>Gets whether a QsiSXTag record relates to a PivotTable view instead of a query table.</summary>
        public bool? QueryTableTagRelatesToPivotTable { get; private set; }

        /// <summary>Gets a stable target name for the QsiSXTag record.</summary>
        public string? QueryTableTagTargetName { get; private set; }

        /// <summary>Gets whether refresh is enabled according to QsiSXTag, when decoded.</summary>
        public bool? QueryTableTagRefreshEnabled { get; private set; }

        /// <summary>Gets whether QsiSXTag marks the associated PivotCache records invalid.</summary>
        public bool? QueryTableTagCacheInvalid { get; private set; }

        /// <summary>Gets whether QsiSXTag marks the PivotTable view as OLAP.</summary>
        public bool? QueryTableTagTensorEx { get; private set; }

        /// <summary>Gets the raw QsiSXTag future option flags.</summary>
        public uint? QueryTableTagFutureOptions { get; private set; }

        /// <summary>Gets the data functionality level that last refreshed the QsiSXTag target.</summary>
        public byte? QueryTableTagLastUpdatedVersion { get; private set; }

        /// <summary>Gets the minimum data functionality level required to refresh the QsiSXTag target.</summary>
        public byte? QueryTableTagUpdatableMinimumVersion { get; private set; }

        /// <summary>Gets the QsiSXTag name offset marker byte.</summary>
        public byte? QueryTableTagNameOffsetMarker { get; private set; }

        /// <summary>Gets the query table or PivotTable view name carried by QsiSXTag.</summary>
        public string? QueryTableTagName { get; private set; }

        /// <summary>Gets the trailing unused QsiSXTag field value.</summary>
        public ushort? QueryTableTagUnused { get; private set; }

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

        internal void SetQueryTableTag(
            ushort futureRecordType,
            ushort futureFlags,
            bool relatesToPivotTable,
            bool refreshEnabled,
            bool cacheInvalid,
            bool tensorEx,
            uint futureOptions,
            byte lastUpdatedVersion,
            byte updatableMinimumVersion,
            byte nameOffsetMarker,
            string name,
            ushort unused) {
            QueryTableTagFutureRecordType = futureRecordType;
            QueryTableTagFutureFlags = futureFlags;
            QueryTableTagRelatesToPivotTable = relatesToPivotTable;
            QueryTableTagTargetName = relatesToPivotTable ? "PivotTable" : "QueryTable";
            QueryTableTagRefreshEnabled = refreshEnabled;
            QueryTableTagCacheInvalid = cacheInvalid;
            QueryTableTagTensorEx = tensorEx;
            QueryTableTagFutureOptions = futureOptions;
            QueryTableTagLastUpdatedVersion = lastUpdatedVersion;
            QueryTableTagUpdatableMinimumVersion = updatableMinimumVersion;
            QueryTableTagNameOffsetMarker = nameOffsetMarker;
            QueryTableTagName = name ?? throw new ArgumentNullException(nameof(name));
            QueryTableTagUnused = unused;
        }

        internal void SetView(
            ushort firstRow,
            ushort lastRow,
            ushort firstColumn,
            ushort lastColumn,
            ushort firstHeaderRow,
            ushort firstDataRow,
            ushort firstDataColumn,
            short cacheIndex,
            ushort dataAxis,
            short dataPosition,
            short fieldCount,
            ushort rowFieldCount,
            ushort columnFieldCount,
            ushort pageFieldCount,
            short dataFieldCount,
            ushort rowLineCount,
            ushort columnLineCount,
            ushort flags,
            ushort autoFormatId,
            string tableName,
            string dataName) {
            ViewFirstRow = firstRow;
            ViewLastRow = lastRow;
            ViewFirstColumn = firstColumn;
            ViewLastColumn = lastColumn;
            ViewRange = FormatRange(firstRow, lastRow, firstColumn, lastColumn);
            ViewFirstHeaderRow = firstHeaderRow;
            ViewFirstDataRow = firstDataRow;
            ViewFirstDataColumn = firstDataColumn;
            ViewCacheIndex = cacheIndex;
            ViewDataAxisName = GetAxisName(dataAxis);
            ViewDataPosition = dataPosition;
            ViewFieldCount = fieldCount;
            ViewRowFieldCount = rowFieldCount;
            ViewColumnFieldCount = columnFieldCount;
            ViewPageFieldCount = pageFieldCount;
            ViewDataFieldCount = dataFieldCount;
            ViewRowLineCount = rowLineCount;
            ViewColumnLineCount = columnLineCount;
            ViewRowGrandTotals = (flags & 0x0001) != 0;
            ViewColumnGrandTotals = (flags & 0x0002) != 0;
            ViewAutoFormat = (flags & 0x0008) != 0;
            ViewAutoFormatId = autoFormatId;
            ViewTableName = tableName ?? throw new ArgumentNullException(nameof(tableName));
            ViewDataName = dataName ?? throw new ArgumentNullException(nameof(dataName));
        }

        internal void SetField(ushort axis, ushort subtotalFlags, ushort subtotalCount, short itemCount, string? fieldName) {
            FieldAxisName = GetAxisName(axis);
            FieldSubtotalCount = subtotalCount;
            FieldSubtotalFlags = subtotalFlags;
            FieldSubtotalFunctionNames = GetSubtotalFunctionNames(subtotalFlags);
            FieldItemCount = itemCount;
            FieldName = fieldName;
        }

        internal void SetFieldIndexList(IReadOnlyList<short> fieldIndexes) {
            FieldIndexReferences = fieldIndexes ?? throw new ArgumentNullException(nameof(fieldIndexes));
        }

        internal void SetLineItems(IReadOnlyList<LegacyXlsPivotLineItem> lineItems) {
            LineItems = lineItems ?? throw new ArgumentNullException(nameof(lineItems));
        }

        internal void SetPageItems(IReadOnlyList<LegacyXlsPivotPageItem> pageItems) {
            PageItems = pageItems ?? throw new ArgumentNullException(nameof(pageItems));
        }

        internal void SetItem(short itemType, ushort flags, short cacheIndex, string? itemName) {
            ItemType = itemType;
            ItemTypeKind = TryGetItemTypeKind(itemType);
            ItemTypeName = ItemTypeKind?.ToString() ?? $"ItemType:{itemType}";
            ItemHidden = (flags & 0x0001) != 0;
            ItemHideDetail = (flags & 0x0002) != 0;
            ItemFormula = (flags & 0x0008) != 0;
            ItemMissing = (flags & 0x0010) != 0;
            ItemCacheIndex = cacheIndex;
            ItemCacheIndexName = cacheIndex == -1
                ? "NoCacheItem"
                : $"CacheItem:{cacheIndex}";
            ItemName = itemName;
        }

        internal void SetGroupingRange(bool autoStart, bool autoEnd, LegacyXlsPivotGroupingKind groupingKind) {
            AutoStart = autoStart;
            AutoEnd = autoEnd;
            GroupingKind = groupingKind;
        }

        internal void SetCalculatedItemFormula(ushort reserved, short cacheFieldIndex) {
            CalculatedItemFormulaReserved = reserved;
            CalculatedItemFormulaCacheFieldIndex = cacheFieldIndex;
            CalculatedItemFormulaScopeName = cacheFieldIndex == -1
                ? "AllCacheFields"
                : $"CacheField:{cacheFieldIndex}";
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

        internal void SetAdditionalClassNesting(int sequenceIndex, int depthBefore, int depthAfter, string transition) {
            AdditionalSequenceIndex = sequenceIndex;
            AdditionalClassDepthBefore = depthBefore;
            AdditionalClassDepthAfter = depthAfter;
            AdditionalClassTransition = string.IsNullOrWhiteSpace(transition) ? null : transition;
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

        private static LegacyXlsPivotItemType? TryGetItemTypeKind(short value) {
            return value >= 0 && value <= 12 ? (LegacyXlsPivotItemType)value : null;
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

        private static string GetAxisName(ushort value) {
            bool row = (value & 0x0001) != 0;
            bool column = (value & 0x0002) != 0;
            bool page = (value & 0x0004) != 0;
            bool data = (value & 0x0008) != 0;
            if (!row && !column && !page && !data) {
                return "None";
            }

            var names = new List<string>(4);
            if (row) {
                names.Add("Row");
            }

            if (column) {
                names.Add("Column");
            }

            if (page) {
                names.Add("Page");
            }

            if (data) {
                names.Add("Data");
            }

            return string.Join("+", names);
        }

        private static string FormatRange(ushort firstRow, ushort lastRow, ushort firstColumn, ushort lastColumn) {
            string start = A1.CellReference(firstRow + 1, firstColumn + 1);
            string end = A1.CellReference(lastRow + 1, lastColumn + 1);
            return start == end ? start : start + ":" + end;
        }

        private static IReadOnlyList<string> GetSubtotalFunctionNames(ushort value) {
            var names = new List<string>(12);
            AddSubtotalName(names, value, 0x0001, "Default");
            AddSubtotalName(names, value, 0x0002, "Sum");
            AddSubtotalName(names, value, 0x0004, "Count");
            AddSubtotalName(names, value, 0x0008, "Average");
            AddSubtotalName(names, value, 0x0010, "Max");
            AddSubtotalName(names, value, 0x0020, "Min");
            AddSubtotalName(names, value, 0x0040, "Product");
            AddSubtotalName(names, value, 0x0080, "CountNumbers");
            AddSubtotalName(names, value, 0x0100, "StdDev");
            AddSubtotalName(names, value, 0x0200, "StdDevPopulation");
            AddSubtotalName(names, value, 0x0400, "Variance");
            AddSubtotalName(names, value, 0x0800, "VariancePopulation");
            return names;
        }

        private static void AddSubtotalName(List<string> names, ushort value, ushort flag, string name) {
            if ((value & flag) != 0) {
                names.Add(name);
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
