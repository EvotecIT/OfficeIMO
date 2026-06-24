namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a preserve-only chart BIFF record discovered during legacy XLS import.
    /// </summary>
    public sealed class LegacyXlsChartRecord {
        /// <summary>
        /// Creates chart BIFF record metadata.
        /// </summary>
        public LegacyXlsChartRecord(
            LegacyXlsChartRecordKind kind,
            string recordName,
            string? sheetName,
            int recordOffset,
            ushort recordType,
            int payloadLength,
            int? sequenceIndex = null,
            int? containerDepthBefore = null,
            int? containerDepthAfter = null,
            string? containerTransition = null,
            string? chartTypeName = null,
            int? chartX = null,
            int? chartY = null,
            int? chartWidth = null,
            int? chartHeight = null,
            LegacyXlsChartGroupOptions? chartGroupOptions = null,
            ushort? axisType = null,
            string? axisTypeName = null,
            ushort? axesUsedCount = null,
            LegacyXlsChartCategorySeriesRange? categorySeriesRange = null,
            LegacyXlsChartCategoryLabelOptions? categoryLabelOptions = null,
            LegacyXlsChartAxisLineFormat? axisLineFormat = null,
            ushort? seriesCategoryDataType = null,
            string? seriesCategoryDataTypeName = null,
            ushort? seriesValueDataType = null,
            string? seriesValueDataTypeName = null,
            ushort? seriesCategoryCount = null,
            ushort? seriesValueCount = null,
            ushort? seriesBubbleSizeDataType = null,
            string? seriesBubbleSizeDataTypeName = null,
            ushort? seriesBubbleSizeCount = null,
            LegacyXlsChartSeriesChartGroupReference? seriesChartGroupReference = null,
            LegacyXlsChartPivotViewReference? pivotViewReference = null,
            ushort? seriesDataCacheIndex = null,
            string? seriesDataCacheIndexName = null,
            ushort? dataFormatPointIndex = null,
            ushort? dataFormatSeriesIndex = null,
            ushort? dataFormatOrder = null,
            string? dataFormatTarget = null,
            ushort? numberFormatId = null,
            ushort? fontIndex = null,
            LegacyXlsChartLineFormat? lineFormat = null,
            LegacyXlsChartAreaFormat? areaFormat = null,
            LegacyXlsChartMarkerFormat? markerFormat = null,
            LegacyXlsChartPieFormat? pieFormat = null,
            LegacyXlsChartAttachedLabel? attachedLabel = null,
            ushort? defaultTextId = null,
            string? defaultTextTargetName = null,
            LegacyXlsChartText? text = null,
            LegacyXlsChartObjectLink? objectLink = null,
            LegacyXlsChartLegend? legend = null,
            LegacyXlsChartTick? tick = null,
            LegacyXlsChartPosition? position = null,
            LegacyXlsChartDataSource? dataSource = null,
            LegacyXlsChartFrame? frame = null,
            LegacyXlsChartPlotGrowth? plotGrowth = null,
            LegacyXlsChartDataTableOptions? dataTableOptions = null,
            LegacyXlsChartSheetProperties? sheetProperties = null,
            LegacyXlsChartValueRange? valueRange = null,
            LegacyXlsChartBarOptions? barOptions = null,
            LegacyXlsChart3DBarShapeOptions? threeDimensionalBarShapeOptions = null,
            LegacyXlsChartScatterOptions? scatterOptions = null,
            LegacyXlsChartFontBasisOptions? fontBasisOptions = null,
            LegacyXlsChartLayout12? layout12 = null,
            LegacyXlsChartFutureRecordInfo? futureRecordInfo = null,
            LegacyXlsChartXmlTokenChain? xmlTokenChain = null,
            LegacyXlsChartPlotAreaLayout12? plotAreaLayout12 = null,
            LegacyXlsChartFutureBlock? futureBlock = null,
            LegacyXlsChartUnits? units = null,
            LegacyXlsChartSeriesList? seriesList = null,
            LegacyXlsChartSeriesFormat? seriesFormat = null) {
            if (payloadLength < 0) {
                throw new ArgumentOutOfRangeException(nameof(payloadLength));
            }

            Kind = kind;
            RecordName = recordName ?? throw new ArgumentNullException(nameof(recordName));
            SheetName = sheetName;
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
            SequenceIndex = sequenceIndex;
            ContainerDepthBefore = containerDepthBefore;
            ContainerDepthAfter = containerDepthAfter;
            ContainerTransition = string.IsNullOrWhiteSpace(containerTransition) ? null : containerTransition;
            ChartTypeName = string.IsNullOrWhiteSpace(chartTypeName) ? null : chartTypeName;
            ChartX = chartX;
            ChartY = chartY;
            ChartWidth = chartWidth;
            ChartHeight = chartHeight;
            ChartGroupOptions = chartGroupOptions;
            AxisType = axisType;
            AxisTypeName = string.IsNullOrWhiteSpace(axisTypeName) ? null : axisTypeName;
            AxesUsedCount = axesUsedCount;
            CategorySeriesRange = categorySeriesRange;
            CategoryLabelOptions = categoryLabelOptions;
            AxisLineFormat = axisLineFormat;
            SeriesCategoryDataType = seriesCategoryDataType;
            SeriesCategoryDataTypeName = string.IsNullOrWhiteSpace(seriesCategoryDataTypeName) ? null : seriesCategoryDataTypeName;
            SeriesValueDataType = seriesValueDataType;
            SeriesValueDataTypeName = string.IsNullOrWhiteSpace(seriesValueDataTypeName) ? null : seriesValueDataTypeName;
            SeriesCategoryCount = seriesCategoryCount;
            SeriesValueCount = seriesValueCount;
            SeriesBubbleSizeDataType = seriesBubbleSizeDataType;
            SeriesBubbleSizeDataTypeName = string.IsNullOrWhiteSpace(seriesBubbleSizeDataTypeName) ? null : seriesBubbleSizeDataTypeName;
            SeriesBubbleSizeCount = seriesBubbleSizeCount;
            SeriesChartGroupReference = seriesChartGroupReference;
            PivotViewReference = pivotViewReference;
            SeriesDataCacheIndex = seriesDataCacheIndex;
            SeriesDataCacheIndexName = string.IsNullOrWhiteSpace(seriesDataCacheIndexName) ? null : seriesDataCacheIndexName;
            DataFormatPointIndex = dataFormatPointIndex;
            DataFormatSeriesIndex = dataFormatSeriesIndex;
            DataFormatOrder = dataFormatOrder;
            DataFormatTarget = string.IsNullOrWhiteSpace(dataFormatTarget) ? null : dataFormatTarget;
            NumberFormatId = numberFormatId;
            FontIndex = fontIndex;
            LineFormat = lineFormat;
            AreaFormat = areaFormat;
            MarkerFormat = markerFormat;
            PieFormat = pieFormat;
            AttachedLabel = attachedLabel;
            DefaultTextId = defaultTextId;
            DefaultTextTargetName = string.IsNullOrWhiteSpace(defaultTextTargetName) ? null : defaultTextTargetName;
            Text = text;
            ObjectLink = objectLink;
            Legend = legend;
            Tick = tick;
            Position = position;
            DataSource = dataSource;
            Frame = frame;
            PlotGrowth = plotGrowth;
            DataTableOptions = dataTableOptions;
            SheetProperties = sheetProperties;
            ValueRange = valueRange;
            BarOptions = barOptions;
            ThreeDimensionalBarShapeOptions = threeDimensionalBarShapeOptions;
            ScatterOptions = scatterOptions;
            FontBasisOptions = fontBasisOptions;
            Layout12 = layout12;
            FutureRecordInfo = futureRecordInfo;
            XmlTokenChain = xmlTokenChain;
            PlotAreaLayout12 = plotAreaLayout12;
            FutureBlock = futureBlock;
            Units = units;
            SeriesList = seriesList;
            SeriesFormat = seriesFormat;
        }

        /// <summary>Gets the shallow chart record category.</summary>
        public LegacyXlsChartRecordKind Kind { get; }

        /// <summary>Gets the BIFF record name.</summary>
        public string RecordName { get; }

        /// <summary>Gets the worksheet or chart sheet name associated with the record, when known.</summary>
        public string? SheetName { get; }

        /// <summary>Gets the byte offset of the BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type identifier.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the BIFF record payload length in bytes.</summary>
        public int PayloadLength { get; }

        /// <summary>Gets the one-based chart record sequence index within the scanned chart scope.</summary>
        public int? SequenceIndex { get; }

        /// <summary>Gets the chart container nesting depth before this record is applied.</summary>
        public int? ContainerDepthBefore { get; }

        /// <summary>Gets the chart container nesting depth after this record is applied.</summary>
        public int? ContainerDepthAfter { get; }

        /// <summary>Gets the shallow nesting transition represented by this chart record.</summary>
        public string? ContainerTransition { get; }

        /// <summary>Gets the decoded chart family name for BIFF chart-type records, when available.</summary>
        public string? ChartTypeName { get; }

        /// <summary>Gets the decoded chart rectangle X coordinate from Chart records, when present.</summary>
        public int? ChartX { get; }

        /// <summary>Gets the decoded chart rectangle Y coordinate from Chart records, when present.</summary>
        public int? ChartY { get; }

        /// <summary>Gets the decoded chart rectangle width from Chart records, when present.</summary>
        public int? ChartWidth { get; }

        /// <summary>Gets the decoded chart rectangle height from Chart records, when present.</summary>
        public int? ChartHeight { get; }

        /// <summary>Gets decoded chart-group options from ChartFormat records, when present.</summary>
        public LegacyXlsChartGroupOptions? ChartGroupOptions { get; }

        /// <summary>Gets the raw axis type from Axis records, when present.</summary>
        public ushort? AxisType { get; }

        /// <summary>Gets the decoded axis type name from Axis records, when present.</summary>
        public string? AxisTypeName { get; }

        /// <summary>Gets the decoded axis group count from AxesUsed records, when present.</summary>
        public ushort? AxesUsedCount { get; }

        /// <summary>Gets decoded category, date, or series axis range metadata from CatSerRange records, when present.</summary>
        public LegacyXlsChartCategorySeriesRange? CategorySeriesRange { get; }

        /// <summary>Gets decoded axis-label metadata from CatLab records, when present.</summary>
        public LegacyXlsChartCategoryLabelOptions? CategoryLabelOptions { get; }

        /// <summary>Gets decoded axis-line formatting target metadata from AxisLineFormat records, when present.</summary>
        public LegacyXlsChartAxisLineFormat? AxisLineFormat { get; }

        /// <summary>Gets the raw category data type from Series records, when present.</summary>
        public ushort? SeriesCategoryDataType { get; }

        /// <summary>Gets the decoded category data type name from Series records, when present.</summary>
        public string? SeriesCategoryDataTypeName { get; }

        /// <summary>Gets the raw value data type from Series records, when present.</summary>
        public ushort? SeriesValueDataType { get; }

        /// <summary>Gets the decoded value data type name from Series records, when present.</summary>
        public string? SeriesValueDataTypeName { get; }

        /// <summary>Gets the category or horizontal value count from Series records, when present.</summary>
        public ushort? SeriesCategoryCount { get; }

        /// <summary>Gets the value or vertical value count from Series records, when present.</summary>
        public ushort? SeriesValueCount { get; }

        /// <summary>Gets the raw bubble-size data type from Series records, when present.</summary>
        public ushort? SeriesBubbleSizeDataType { get; }

        /// <summary>Gets the decoded bubble-size data type name from Series records, when present.</summary>
        public string? SeriesBubbleSizeDataTypeName { get; }

        /// <summary>Gets the bubble-size value count from Series records, when present.</summary>
        public ushort? SeriesBubbleSizeCount { get; }

        /// <summary>Gets decoded series-to-chart-group linkage from SerToCrt records, when present.</summary>
        public LegacyXlsChartSeriesChartGroupReference? SeriesChartGroupReference { get; }

        /// <summary>Gets decoded SeriesList membership metadata, when present.</summary>
        public LegacyXlsChartSeriesList? SeriesList { get; }

        /// <summary>Gets decoded PivotTable-view range metadata from SBaseRef records, when present.</summary>
        public LegacyXlsChartPivotViewReference? PivotViewReference { get; }

        /// <summary>Gets the raw data-cache sequence index from SIIndex records, when present.</summary>
        public ushort? SeriesDataCacheIndex { get; }

        /// <summary>Gets the decoded data-cache sequence name from SIIndex records, when present.</summary>
        public string? SeriesDataCacheIndexName { get; }

        /// <summary>Gets the raw data-point index from DataFormat records, when present.</summary>
        public ushort? DataFormatPointIndex { get; }

        /// <summary>Gets the raw series index from DataFormat records, when present.</summary>
        public ushort? DataFormatSeriesIndex { get; }

        /// <summary>Gets the raw series order or format index from DataFormat records, when present.</summary>
        public ushort? DataFormatOrder { get; }

        /// <summary>Gets whether a DataFormat record targets a whole series or a point, when present.</summary>
        public string? DataFormatTarget { get; }

        /// <summary>Gets the raw axis number format identifier from IFmtRecord records, when present.</summary>
        public ushort? NumberFormatId { get; }

        /// <summary>Gets the raw font index from FontX records, when present.</summary>
        public ushort? FontIndex { get; }

        /// <summary>Gets decoded line-format metadata from LineFormat records, when present.</summary>
        public LegacyXlsChartLineFormat? LineFormat { get; }

        /// <summary>Gets decoded fill-format metadata from AreaFormat records, when present.</summary>
        public LegacyXlsChartAreaFormat? AreaFormat { get; }

        /// <summary>Gets decoded marker-format metadata from MarkerFormat records, when present.</summary>
        public LegacyXlsChartMarkerFormat? MarkerFormat { get; }

        /// <summary>Gets decoded pie or doughnut explosion metadata from PieFormat records, when present.</summary>
        public LegacyXlsChartPieFormat? PieFormat { get; }

        /// <summary>Gets decoded series-format flags from SerFmt records, when present.</summary>
        public LegacyXlsChartSeriesFormat? SeriesFormat { get; }

        /// <summary>Gets decoded data-label display metadata from AttachedLabel records, when present.</summary>
        public LegacyXlsChartAttachedLabel? AttachedLabel { get; }

        /// <summary>Gets the raw DefaultText target identifier, when present.</summary>
        public ushort? DefaultTextId { get; }

        /// <summary>Gets the decoded DefaultText target name, when present.</summary>
        public string? DefaultTextTargetName { get; }

        /// <summary>Gets decoded text metadata from Text records, when present.</summary>
        public LegacyXlsChartText? Text { get; }

        /// <summary>Gets decoded linked-object metadata from ObjectLink records, when present.</summary>
        public LegacyXlsChartObjectLink? ObjectLink { get; }

        /// <summary>Gets decoded legend metadata from Legend records, when present.</summary>
        public LegacyXlsChartLegend? Legend { get; }

        /// <summary>Gets decoded axis tick metadata from Tick records, when present.</summary>
        public LegacyXlsChartTick? Tick { get; }

        /// <summary>Gets decoded position metadata from Pos records, when present.</summary>
        public LegacyXlsChartPosition? Position { get; }

        /// <summary>Gets decoded data-source metadata from BRAI records, when present.</summary>
        public LegacyXlsChartDataSource? DataSource { get; }

        /// <summary>Gets decoded frame metadata from Frame records, when present.</summary>
        public LegacyXlsChartFrame? Frame { get; }

        /// <summary>Gets decoded font-scaling metadata from PlotGrowth records, when present.</summary>
        public LegacyXlsChartPlotGrowth? PlotGrowth { get; }

        /// <summary>Gets decoded chart data-table display options from Dat records, when present.</summary>
        public LegacyXlsChartDataTableOptions? DataTableOptions { get; }

        /// <summary>Gets decoded chart sheet properties from ShtProps records, when present.</summary>
        public LegacyXlsChartSheetProperties? SheetProperties { get; }

        /// <summary>Gets decoded value-axis scale metadata from ValueRange records, when present.</summary>
        public LegacyXlsChartValueRange? ValueRange { get; }

        /// <summary>Gets decoded bar or column chart group options from Bar records, when present.</summary>
        public LegacyXlsChartBarOptions? BarOptions { get; }

        /// <summary>Gets decoded 3-D bar or column data-point shape options from Chart3DBarShape records, when present.</summary>
        public LegacyXlsChart3DBarShapeOptions? ThreeDimensionalBarShapeOptions { get; }

        /// <summary>Gets decoded scatter or bubble chart group options from Scatter records, when present.</summary>
        public LegacyXlsChartScatterOptions? ScatterOptions { get; }

        /// <summary>Gets decoded chart font-scaling metadata from Fbi records, when present.</summary>
        public LegacyXlsChartFontBasisOptions? FontBasisOptions { get; }

        /// <summary>Gets decoded chart layout metadata from CrtLayout12 records, when present.</summary>
        public LegacyXlsChartLayout12? Layout12 { get; }

        /// <summary>Gets decoded chart future-record range metadata from ChartFrtInfo records, when present.</summary>
        public LegacyXlsChartFutureRecordInfo? FutureRecordInfo { get; }

        /// <summary>Gets decoded chart XML token-chain metadata from CrtMlFrt records, when present.</summary>
        public LegacyXlsChartXmlTokenChain? XmlTokenChain { get; }

        /// <summary>Gets decoded plot-area layout metadata from CrtLayout12A records, when present.</summary>
        public LegacyXlsChartPlotAreaLayout12? PlotAreaLayout12 { get; }

        /// <summary>Gets decoded future-record block scope metadata from StartBlock and EndBlock records, when present.</summary>
        public LegacyXlsChartFutureBlock? FutureBlock { get; }

        /// <summary>Gets decoded preserve-only Units metadata, when present.</summary>
        public LegacyXlsChartUnits? Units { get; }
    }
}
