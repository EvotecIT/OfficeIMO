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

        private static LegacyXlsPivotAggregationFunction? TryGetAggregationFunctionKind(short value) {
            return value >= 0 && value <= 10 ? (LegacyXlsPivotAggregationFunction)value : null;
        }

        private static LegacyXlsPivotDisplayCalculation? TryGetDisplayCalculationKind(short value) {
            return value >= 0 && value <= 8 ? (LegacyXlsPivotDisplayCalculation)value : null;
        }
    }
}
