namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a worksheet parsed from a legacy XLS workbook stream.
    /// </summary>
    public sealed class LegacyXlsWorksheet {
        private readonly List<LegacyXlsCell> _cells = new();
        private readonly List<LegacyXlsColumnLayout> _columns = new();
        private readonly List<LegacyXlsCellWatch> _cellWatches = new();
        private readonly List<LegacyXlsComment> _comments = new();
        private readonly List<LegacyXlsConditionalFormattingExtensionRecord> _conditionalFormattingExtensions = new();
        private readonly List<LegacyXlsConditionalFormatting> _conditionalFormattings = new();
        private readonly List<LegacyXlsDataValidationCollectionRecord> _dataValidationCollections = new();
        private readonly List<LegacyXlsDataValidation> _dataValidations = new();
        private readonly List<LegacyXlsArrayFormulaRecord> _arrayFormulaRecords = new();
        private readonly List<LegacyXlsAutoFilterCriteria> _autoFilterCriteria = new();
        private readonly List<LegacyXlsPageBreak> _columnPageBreaks = new();
        private readonly List<LegacyXlsHyperlink> _hyperlinks = new();
        private readonly List<LegacyXlsIgnoredError> _ignoredErrors = new();
        private readonly List<LegacyXlsMergedRange> _mergedRanges = new();
        private readonly List<LegacyXlsWorksheetMetadataRecord> _metadataRecords = new();
        private readonly List<LegacyXlsSheetFutureMetadataRecord> _futureMetadataRecords = new();
        private readonly List<LegacyXlsProtectedRange> _protectedRanges = new();
        private readonly List<LegacyXlsPageBreak> _rowPageBreaks = new();
        private readonly List<LegacyXlsScenario> _scenarios = new();
        private readonly List<LegacyXlsRowLayout> _rows = new();
        private readonly List<LegacyXlsSelection> _selections = new();
        private readonly List<LegacyXlsTableDefinition> _tableDefinitions = new();
        private readonly List<LegacyXlsWorksheetWindowView> _windowViews = new();

        /// <summary>
        /// Creates a parsed legacy XLS worksheet.
        /// </summary>
        /// <param name="name">Worksheet name.</param>
        /// <param name="streamOffset">Byte offset of the worksheet substream in the BIFF workbook stream.</param>
        /// <param name="visibility">Legacy worksheet visibility flag.</param>
        /// <param name="sheetType">Legacy sheet type flag.</param>
        public LegacyXlsWorksheet(string name, int streamOffset, byte visibility, byte sheetType) {
            Name = name;
            StreamOffset = streamOffset;
            Visibility = visibility;
            SheetType = sheetType;
        }

        /// <summary>
        /// Gets the worksheet name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the byte offset of the worksheet substream in the BIFF workbook stream.
        /// </summary>
        public int StreamOffset { get; }

        /// <summary>
        /// Gets the legacy visibility flag.
        /// </summary>
        public byte Visibility { get; }

        /// <summary>
        /// Gets the decoded sheet visibility state, when the BoundSheet value is recognized.
        /// </summary>
        public LegacyXlsSheetVisibility? VisibilityKind => LegacyXlsSheetVisibilityDecoder.ToKind(Visibility);

        /// <summary>
        /// Gets the decoded sheet visibility state name, or a hexadecimal fallback for unknown values.
        /// </summary>
        public string VisibilityName => LegacyXlsSheetVisibilityDecoder.ToName(Visibility);

        /// <summary>
        /// Gets the legacy sheet type flag.
        /// </summary>
        public byte SheetType { get; }

        /// <summary>
        /// Gets the sheet object name used by the VBA project, when specified.
        /// </summary>
        public string? CodeName { get; private set; }

        /// <summary>
        /// Gets the parsed cells for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsCell> Cells => _cells;

        /// <summary>
        /// Gets parsed column layout metadata for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsColumnLayout> Columns => _columns;

        /// <summary>
        /// Gets watched-cell metadata parsed for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsCellWatch> CellWatches => _cellWatches;

        /// <summary>
        /// Gets parsed cell comments for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsComment> Comments => _comments;

        /// <summary>
        /// Gets parsed conditional formatting rules for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsConditionalFormatting> ConditionalFormattings => _conditionalFormattings;

        /// <summary>
        /// Gets preserve-only conditional-formatting extension records for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsConditionalFormattingExtensionRecord> ConditionalFormattingExtensions => _conditionalFormattingExtensions;

        /// <summary>
        /// Gets parsed data validation rules for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsDataValidation> DataValidations => _dataValidations;

        /// <summary>
        /// Gets parsed DVal collection headers for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsDataValidationCollectionRecord> DataValidationCollections => _dataValidationCollections;

        /// <summary>
        /// Gets worksheet-level data-consolidation settings parsed from a BIFF DCon record.
        /// </summary>
        public LegacyXlsDataConsolidationSettings? DataConsolidationSettings { get; private set; }

        /// <summary>
        /// Gets preserve/projection metadata for Array formula records discovered on this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsArrayFormulaRecord> ArrayFormulaRecords => _arrayFormulaRecords;

        /// <summary>
        /// Gets parsed AutoFilter criteria for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsAutoFilterCriteria> AutoFilterCriteria => _autoFilterCriteria;

        /// <summary>
        /// Gets the declared AutoFilter drop-down count, when present.
        /// </summary>
        public ushort? AutoFilterDropDownCount { get; private set; }

        /// <summary>
        /// Gets explicit manual column page breaks parsed for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsPageBreak> ColumnPageBreaks => _columnPageBreaks;

        /// <summary>
        /// Gets worksheet-level phonetic display defaults parsed from a BIFF PhoneticInfo record.
        /// </summary>
        public LegacyXlsPhoneticSettings? PhoneticSettings { get; private set; }

        /// <summary>
        /// Gets parsed hyperlinks for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsHyperlink> Hyperlinks => _hyperlinks;

        /// <summary>
        /// Gets ignored formula error metadata parsed from BIFF8 shared-feature records.
        /// </summary>
        public IReadOnlyList<LegacyXlsIgnoredError> IgnoredErrors => _ignoredErrors;

        /// <summary>
        /// Gets protected ranges parsed from BIFF8 worksheet shared-feature records.
        /// </summary>
        public IReadOnlyList<LegacyXlsProtectedRange> ProtectedRanges => _protectedRanges;

        /// <summary>
        /// Gets parsed merged ranges for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsMergedRange> MergedRanges => _mergedRanges;

        /// <summary>
        /// Gets decoded worksheet metadata source records.
        /// </summary>
        public IReadOnlyList<LegacyXlsWorksheetMetadataRecord> MetadataRecords => _metadataRecords;

        /// <summary>
        /// Gets preserve-only extended metadata records decoded from this worksheet or chart-sheet substream.
        /// </summary>
        public IReadOnlyList<LegacyXlsSheetFutureMetadataRecord> FutureMetadataRecords => _futureMetadataRecords;

        /// <summary>
        /// Gets explicit manual row page breaks parsed for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsPageBreak> RowPageBreaks => _rowPageBreaks;

        /// <summary>
        /// Gets worksheet scenarios parsed from SCENARIO records.
        /// </summary>
        public IReadOnlyList<LegacyXlsScenario> Scenarios => _scenarios;

        /// <summary>
        /// Gets parsed row layout metadata for this worksheet.
        /// </summary>
        public IReadOnlyList<LegacyXlsRowLayout> Rows => _rows;

        /// <summary>
        /// Gets parsed worksheet selections.
        /// </summary>
        public IReadOnlyList<LegacyXlsSelection> Selections => _selections;

        /// <summary>
        /// Gets worksheet table/list definitions parsed from BIFF shared-feature records.
        /// </summary>
        public IReadOnlyList<LegacyXlsTableDefinition> TableDefinitions => _tableDefinitions;

        /// <summary>
        /// Gets worksheet-window view records parsed from BIFF Window2 records.
        /// </summary>
        public IReadOnlyList<LegacyXlsWorksheetWindowView> WindowViews => _windowViews;

        /// <summary>
        /// Gets parsed worksheet sort dialog metadata, when present.
        /// </summary>
        public LegacyXlsSortSettings? SortSettings { get; private set; }

        /// <summary>
        /// Gets scenario-manager metadata parsed from a ScenMan record.
        /// </summary>
        public LegacyXlsScenarioManager? ScenarioManager { get; private set; }

        /// <summary>
        /// Gets parsed frozen pane metadata for this worksheet.
        /// </summary>
        public LegacyXlsFreezePane? FreezePane { get; private set; }

        /// <summary>
        /// Gets parsed non-frozen split pane metadata for this worksheet.
        /// </summary>
        public LegacyXlsSplitPane? SplitPane { get; private set; }

        /// <summary>
        /// Gets whether the frozen pane Window2 record requests the BIFF frozen-without-split state.
        /// </summary>
        public bool? FrozenWithoutSplit { get; private set; }

        /// <summary>
        /// Gets parsed worksheet protection metadata.
        /// </summary>
        public LegacyXlsWorksheetProtection? Protection { get; private set; }

        /// <summary>
        /// Gets the used-range bounds declared by the legacy worksheet DIMENSIONS record, when present.
        /// </summary>
        public LegacyXlsWorksheetDimension? DeclaredUsedRange { get; private set; }

        /// <summary>
        /// Gets parsed worksheet page setup metadata.
        /// </summary>
        public LegacyXlsPageSetup? PageSetup { get; private set; }

        /// <summary>
        /// Gets parsed row-block lookup metadata from the BIFF Index record, when present.
        /// </summary>
        public LegacyXlsWorksheetIndex? RowBlockIndex { get; private set; }

        /// <summary>
        /// Gets whether automatic page breaks should be visible, when present.
        /// </summary>
        public bool? AutomaticPageBreaksVisible { get; private set; }

        /// <summary>
        /// Gets whether outline styles should be applied automatically, when present.
        /// </summary>
        public bool? ApplyOutlineStyles { get; private set; }

        /// <summary>
        /// Gets whether row summaries should appear below detail rows, when present.
        /// </summary>
        public bool? SummaryRowsBelow { get; private set; }

        /// <summary>
        /// Gets whether column summaries should appear to the right in left-to-right sheets, when present.
        /// </summary>
        public bool? SummaryColumnsRightWhenLeftToRight { get; private set; }

        /// <summary>
        /// Gets whether horizontal scrolling should be synchronized with another sheet, when present.
        /// </summary>
        public bool? SynchronizedHorizontalScrolling { get; private set; }

        /// <summary>
        /// Gets whether vertical scrolling should be synchronized with another sheet, when present.
        /// </summary>
        public bool? SynchronizedVerticalScrolling { get; private set; }

        /// <summary>
        /// Gets whether transition formula evaluation is enabled, when present.
        /// </summary>
        public bool? TransitionFormulaEvaluation { get; private set; }

        /// <summary>
        /// Gets whether transition formula entry is enabled, when present.
        /// </summary>
        public bool? TransitionFormulaEntry { get; private set; }

        /// <summary>
        /// Gets the maximum row outline level, when present.
        /// </summary>
        public byte? RowOutlineLevel { get; private set; }

        /// <summary>
        /// Gets the maximum column outline level, when present.
        /// </summary>
        public byte? ColumnOutlineLevel { get; private set; }

        /// <summary>
        /// Gets the legacy GridSet flag, when present.
        /// </summary>
        public bool? GridSet { get; private set; }

        /// <summary>
        /// Gets whether formulas on this worksheet should be recalculated on load.
        /// </summary>
        public bool? FullCalculationOnLoad { get; private set; }

        /// <summary>
        /// Gets the worksheet view zoom scale percentage, when present.
        /// </summary>
        public uint? ZoomScale { get; private set; }

        /// <summary>
        /// Gets the normal-view zoom scale percentage, when specified by the legacy window metadata.
        /// </summary>
        public uint? ZoomScaleNormal { get; private set; }

        /// <summary>
        /// Gets the default row height in points, when specified by the legacy worksheet metadata.
        /// </summary>
        public double? DefaultRowHeight { get; private set; }

        /// <summary>
        /// Gets the default column width in character units, when specified by the legacy worksheet metadata.
        /// </summary>
        public double? DefaultColumnWidth { get; private set; }

        /// <summary>
        /// Gets whether empty rows are hidden by default, when specified by the legacy worksheet metadata.
        /// </summary>
        public bool DefaultRowsHidden { get; private set; }

        /// <summary>
        /// Gets whether worksheet view gridlines should be shown, when specified by the legacy window metadata.
        /// </summary>
        public bool? ShowGridLines { get; private set; }

        /// <summary>
        /// Gets whether cell formulas should be shown in the worksheet view, when specified by the legacy window metadata.
        /// </summary>
        public bool? ShowFormulas { get; private set; }

        /// <summary>
        /// Gets whether worksheet row and column headings should be shown, when specified by the legacy window metadata.
        /// </summary>
        public bool? ShowRowColumnHeadings { get; private set; }

        /// <summary>
        /// Gets whether zero values should be shown in the worksheet view, when specified by the legacy window metadata.
        /// </summary>
        public bool? ShowZeroValues { get; private set; }

        /// <summary>
        /// Gets whether the worksheet view should be displayed from right to left, when specified by the legacy window metadata.
        /// </summary>
        public bool? RightToLeft { get; private set; }

        /// <summary>
        /// Gets whether the worksheet uses the default gridline color, when specified by the legacy window metadata.
        /// </summary>
        public bool? DefaultGridColor { get; private set; }

        /// <summary>
        /// Gets the worksheet view gridline color index, when specified by the legacy window metadata.
        /// </summary>
        public ushort? GridLineColorIndex { get; private set; }

        /// <summary>
        /// Gets whether outline symbols should be shown in the worksheet view, when specified by the legacy window metadata.
        /// </summary>
        public bool? ShowOutlineSymbols { get; private set; }

        /// <summary>
        /// Gets whether the worksheet tab is selected, when specified by the legacy window metadata.
        /// </summary>
        public bool? TabSelected { get; private set; }

        /// <summary>
        /// Gets whether the worksheet is displayed in Page Break Preview view, when specified by the legacy window metadata.
        /// </summary>
        public bool? PageBreakPreview { get; private set; }

        /// <summary>
        /// Gets whether the worksheet is displayed in Page Layout view, when specified by a PLV future record.
        /// </summary>
        public bool? PageLayoutView { get; private set; }

        /// <summary>
        /// Gets the Page Layout view zoom scale, when specified by a PLV future record.
        /// </summary>
        public uint? PageLayoutZoomScale { get; private set; }

        /// <summary>
        /// Gets the BIFF color index used for the worksheet tab color, when specified by a SheetExt record.
        /// </summary>
        public ushort? TabColorIndex { get; private set; }

        /// <summary>
        /// Gets the zero-based first visible row in the worksheet window, when specified by the legacy window metadata.
        /// </summary>
        public int? FirstVisibleRow { get; private set; }

        /// <summary>
        /// Gets the zero-based first visible column in the worksheet window, when specified by the legacy window metadata.
        /// </summary>
        public int? FirstVisibleColumn { get; private set; }

        internal void AddCell(LegacyXlsCell cell) {
            _cells.Add(cell);
        }

        internal bool TryReplaceFormulaText(int row, int column, string formulaText) {
            for (int i = _cells.Count - 1; i >= 0; i--) {
                LegacyXlsCell cell = _cells[i];
                if (cell.Row == row && cell.Column == column && cell.IsFormula) {
                    _cells[i] = new LegacyXlsCell(
                        cell.Row,
                        cell.Column,
                        cell.Kind,
                        cell.Value,
                        cell.StyleIndex,
                        isFormula: true,
                        formulaText: formulaText,
                        textFormattingRuns: cell.TextFormattingRuns);
                    return true;
                }
            }

            return false;
        }

        internal void AddColumn(LegacyXlsColumnLayout column) {
            _columns.Add(column);
        }

        internal void AddComment(LegacyXlsComment comment) {
            _comments.Add(comment);
        }

        internal void AddConditionalFormatting(LegacyXlsConditionalFormatting conditionalFormatting) {
            _conditionalFormattings.Add(conditionalFormatting);
        }

        internal void AddConditionalFormattingExtension(LegacyXlsConditionalFormattingExtensionRecord extensionRecord) {
            _conditionalFormattingExtensions.Add(extensionRecord);
        }

        internal void AddDataValidation(LegacyXlsDataValidation validation) {
            _dataValidations.Add(validation);
        }

        internal void AddDataValidationCollection(LegacyXlsDataValidationCollectionRecord collectionRecord) {
            _dataValidationCollections.Add(collectionRecord);
        }

        internal void AddArrayFormulaRecord(LegacyXlsArrayFormulaRecord arrayFormulaRecord) {
            _arrayFormulaRecords.Add(arrayFormulaRecord);
        }

        internal void AddAutoFilterCriteria(LegacyXlsAutoFilterCriteria criteria) {
            _autoFilterCriteria.Add(criteria);
        }

        internal void SetAutoFilterDropDownCount(ushort count) {
            AutoFilterDropDownCount = count;
        }

        internal void SetDataConsolidationSettings(LegacyXlsDataConsolidationSettings settings) {
            DataConsolidationSettings = settings;
        }

        internal void AddColumnPageBreak(LegacyXlsPageBreak pageBreak) {
            _columnPageBreaks.Add(pageBreak);
        }

        internal void AddCellWatch(LegacyXlsCellWatch cellWatch) {
            _cellWatches.Add(cellWatch);
        }

        internal void SetScenarioManager(LegacyXlsScenarioManager scenarioManager) {
            ScenarioManager = scenarioManager;
        }

        internal void AddScenario(LegacyXlsScenario scenario) {
            _scenarios.Add(scenario);
        }

        internal void AddHyperlink(LegacyXlsHyperlink hyperlink) {
            _hyperlinks.Add(hyperlink);
        }

        internal void AddIgnoredError(LegacyXlsIgnoredError ignoredError) {
            _ignoredErrors.Add(ignoredError);
        }

        internal void AddProtectedRange(LegacyXlsProtectedRange protectedRange) {
            _protectedRanges.Add(protectedRange);
        }

        internal bool TrySetHyperlinkTooltip(int startRow, int startColumn, int endRow, int endColumn, string tooltip) {
            for (int i = _hyperlinks.Count - 1; i >= 0; i--) {
                LegacyXlsHyperlink hyperlink = _hyperlinks[i];
                if (hyperlink.StartRow == startRow
                    && hyperlink.StartColumn == startColumn
                    && hyperlink.EndRow == endRow
                    && hyperlink.EndColumn == endColumn) {
                    _hyperlinks[i] = hyperlink.WithTooltip(tooltip);
                    return true;
                }
            }

            return false;
        }

        internal void AddMergedRange(LegacyXlsMergedRange mergedRange) {
            _mergedRanges.Add(mergedRange);
        }

        internal void AddMetadataRecord(LegacyXlsWorksheetMetadataKind kind, int recordOffset, ushort recordType) {
            _metadataRecords.Add(new LegacyXlsWorksheetMetadataRecord(kind, recordOffset, recordType));
        }

        internal void AddFutureMetadataRecord(LegacyXlsSheetFutureMetadataRecord record) {
            _futureMetadataRecords.Add(record);
            AddMetadataRecord(LegacyXlsWorksheetMetadataKind.FutureMetadata, record.RecordOffset, record.RecordType);
        }

        internal void SetCodeName(string? value) {
            CodeName = value;
        }

        internal void AddRowPageBreak(LegacyXlsPageBreak pageBreak) {
            _rowPageBreaks.Add(pageBreak);
        }

        internal void AddRow(LegacyXlsRowLayout row) {
            _rows.Add(row);
        }

        internal void AddSelection(LegacyXlsSelection selection) {
            _selections.Add(selection);
        }

        internal void AddTableDefinition(LegacyXlsTableDefinition tableDefinition) {
            _tableDefinitions.Add(tableDefinition);
        }

        internal bool TryApplyTableDefinitionStyle(
            uint idList,
            string styleName,
            bool showFirstColumn,
            bool showLastColumn,
            bool showRowStripes,
            bool showColumnStripes) {
            LegacyXlsTableDefinition? tableDefinition = _tableDefinitions.LastOrDefault(table => table.IdList == idList);
            if (tableDefinition == null) {
                return false;
            }

            tableDefinition.ApplyStyle(styleName, showFirstColumn, showLastColumn, showRowStripes, showColumnStripes);
            return true;
        }

        internal bool TryApplyTableDefinitionDisplayMetadata(uint idList, string? displayName, string? comment) {
            LegacyXlsTableDefinition? tableDefinition = _tableDefinitions.LastOrDefault(table => table.IdList == idList);
            if (tableDefinition == null) {
                return false;
            }

            tableDefinition.ApplyDisplayMetadata(displayName, comment);
            return true;
        }

        internal bool TryApplyTableDefinitionBlockLevelFormatting(uint idList, LegacyXlsTableBlockLevelFormatting formatting) {
            LegacyXlsTableDefinition? tableDefinition = _tableDefinitions.LastOrDefault(table => table.IdList == idList);
            if (tableDefinition == null) {
                return false;
            }

            tableDefinition.ApplyBlockLevelFormatting(formatting);
            return true;
        }

        internal void AddWindowView(LegacyXlsWorksheetWindowView view) {
            _windowViews.Add(view);
        }

        internal void SetSortSettings(LegacyXlsSortSettings sortSettings) {
            SortSettings = sortSettings;
        }

        internal void SetFreezePane(LegacyXlsFreezePane freezePane) {
            FreezePane = freezePane;
        }

        internal void SetSplitPane(LegacyXlsSplitPane splitPane) {
            SplitPane = splitPane;
        }

        internal void SetFrozenWithoutSplit(bool frozenWithoutSplit) {
            FrozenWithoutSplit = frozenWithoutSplit;
        }

        internal void SetZoomScale(uint zoomScale) {
            ZoomScale = zoomScale;
        }

        internal void SetZoomScaleNormal(uint zoomScale) {
            ZoomScaleNormal = zoomScale;
        }

        internal void SetDefaultRowHeight(double height, bool hidden) {
            DefaultRowHeight = height;
            DefaultRowsHidden = hidden;
        }

        internal void SetDefaultColumnWidth(double width) {
            DefaultColumnWidth = width;
        }

        internal void SetGridLinesVisible(bool visible) {
            ShowGridLines = visible;
        }

        internal void SetFormulasVisible(bool visible) {
            ShowFormulas = visible;
        }

        internal void SetRowColumnHeadingsVisible(bool visible) {
            ShowRowColumnHeadings = visible;
        }

        internal void SetZeroValuesVisible(bool visible) {
            ShowZeroValues = visible;
        }

        internal void SetRightToLeft(bool rightToLeft) {
            RightToLeft = rightToLeft;
        }

        internal void SetDefaultGridColor(bool defaultGridColor) {
            DefaultGridColor = defaultGridColor;
        }

        internal void SetGridLineColorIndex(ushort colorIndex) {
            GridLineColorIndex = colorIndex;
        }

        internal void SetOutlineSymbolsVisible(bool visible) {
            ShowOutlineSymbols = visible;
        }

        internal void SetTabSelected(bool selected) {
            TabSelected = selected;
        }

        internal void SetPageBreakPreview(bool enabled) {
            PageBreakPreview = enabled;
        }

        internal void SetPageLayoutView(bool enabled, uint? zoomScale = null) {
            PageLayoutView = enabled;
            PageLayoutZoomScale = zoomScale;
        }

        internal void SetTabColorIndex(ushort colorIndex) {
            TabColorIndex = colorIndex;
        }

        internal void SetFirstVisibleCell(int row, int column) {
            FirstVisibleRow = row;
            FirstVisibleColumn = column;
        }

        internal void SetDeclaredUsedRange(LegacyXlsWorksheetDimension dimension) {
            DeclaredUsedRange = dimension;
        }

        internal LegacyXlsPageSetup GetOrCreatePageSetup() {
            return PageSetup ??= new LegacyXlsPageSetup();
        }

        internal void SetGridSet(bool gridSet) {
            GridSet = gridSet;
        }

        internal void SetFullCalculationOnLoad(bool fullCalculationOnLoad) {
            FullCalculationOnLoad = fullCalculationOnLoad;
        }

        internal void SetPhoneticSettings(LegacyXlsPhoneticSettings settings) {
            PhoneticSettings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        internal void SetOutlineLevels(byte rowLevel, byte columnLevel) {
            RowOutlineLevel = rowLevel;
            ColumnOutlineLevel = columnLevel;
        }

        internal void SetRowBlockIndex(LegacyXlsWorksheetIndex rowBlockIndex) {
            RowBlockIndex = rowBlockIndex;
        }

        internal void SetSheetOptions(ushort options) {
            AutomaticPageBreaksVisible = (options & 0x0001) != 0;
            ApplyOutlineStyles = (options & 0x0020) != 0;
            SummaryRowsBelow = (options & 0x0040) != 0;
            SummaryColumnsRightWhenLeftToRight = (options & 0x0080) != 0;
            GetOrCreatePageSetup().FitToPage = (options & 0x0100) != 0;
            SynchronizedHorizontalScrolling = (options & 0x1000) != 0;
            SynchronizedVerticalScrolling = (options & 0x2000) != 0;
            TransitionFormulaEvaluation = (options & 0x4000) != 0;
            TransitionFormulaEntry = (options & 0x8000) != 0;
        }

        internal void SetProtection(bool isProtected) {
            Protection = new LegacyXlsWorksheetProtection(
                isProtected,
                Protection?.LegacyPasswordHash,
                Protection?.ProtectObjects,
                Protection?.ProtectScenarios,
                Protection?.Permissions);
        }

        internal void SetProtectionPasswordHash(ushort passwordHash) {
            Protection = (Protection ?? new LegacyXlsWorksheetProtection(isProtected: false)).WithLegacyPasswordHash(passwordHash);
        }

        internal void SetObjectProtection(bool isProtected) {
            Protection = (Protection ?? new LegacyXlsWorksheetProtection(isProtected: false)).WithObjectProtection(isProtected);
        }

        internal void SetScenarioProtection(bool isProtected) {
            Protection = (Protection ?? new LegacyXlsWorksheetProtection(isProtected: false)).WithScenarioProtection(isProtected);
        }

        internal void SetProtectionPermissions(LegacyXlsWorksheetProtectionPermissions permissions) {
            Protection = (Protection ?? new LegacyXlsWorksheetProtection(isProtected: false)).WithPermissions(permissions);
        }
    }
}
