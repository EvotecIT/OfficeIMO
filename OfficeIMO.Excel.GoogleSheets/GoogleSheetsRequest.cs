namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>Updates target spreadsheet locale, time zone, and calculation settings.</summary>
    public sealed class GoogleSheetsUpdateSpreadsheetPropertiesRequest : GoogleSheetsRequest {
        public GoogleSheetsUpdateSpreadsheetPropertiesRequest() : base("updateSpreadsheetProperties") { }
        public string? Locale { get; set; }
        public string? TimeZone { get; set; }
        public GoogleSheetsRecalculationInterval RecalculationInterval { get; set; }
    }

    /// <summary>
    /// Base request emitted by the Google Sheets batch compiler.
    /// </summary>
    public abstract class GoogleSheetsRequest {
        protected GoogleSheetsRequest(string kind) {
            Kind = kind ?? throw new ArgumentNullException(nameof(kind));
        }

        public string Kind { get; }
    }

    /// <summary>
    /// Adds or configures a worksheet in the target spreadsheet.
    /// </summary>
    public sealed class GoogleSheetsAddSheetRequest : GoogleSheetsRequest {
        public GoogleSheetsAddSheetRequest() : base("addSheet") {
        }

        public string SheetName { get; set; } = string.Empty;
        public int SheetIndex { get; set; }
        public bool Hidden { get; set; }
        public bool RightToLeft { get; set; }
        public string? TabColorArgb { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public int FrozenRowCount { get; set; }
        public int FrozenColumnCount { get; set; }
        public bool HideGridlines { get; set; }
    }

    /// <summary>
    /// Writes cell values into a worksheet.
    /// </summary>
    public sealed class GoogleSheetsUpdateCellsRequest : GoogleSheetsRequest {
        private readonly List<GoogleSheetsCellData> _cells = new List<GoogleSheetsCellData>();

        public GoogleSheetsUpdateCellsRequest() : base("updateCells") {
        }

        public string SheetName { get; set; } = string.Empty;
        public IReadOnlyList<GoogleSheetsCellData> Cells => _cells;

        internal void AddCell(GoogleSheetsCellData cell) {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            _cells.Add(cell);
        }
    }

    /// <summary>
    /// Applies one validation rule to an entire worksheet range without materializing every target cell.
    /// </summary>
    public sealed class GoogleSheetsSetDataValidationRequest : GoogleSheetsRequest {
        public GoogleSheetsSetDataValidationRequest() : base("setDataValidation") {
        }

        public string SheetName { get; set; } = string.Empty;
        public string A1Range { get; set; } = string.Empty;
        public int StartRowIndex { get; set; }
        public int EndRowIndexExclusive { get; set; }
        public int StartColumnIndex { get; set; }
        public int EndColumnIndexExclusive { get; set; }
        public GoogleSheetsDataValidationRule Rule { get; set; } = new GoogleSheetsDataValidationRule();
    }

    /// <summary>
    /// Adds a named range to the target spreadsheet.
    /// </summary>
    public sealed class GoogleSheetsAddNamedRangeRequest : GoogleSheetsRequest {
        public GoogleSheetsAddNamedRangeRequest() : base("addNamedRange") {
        }

        public string Name { get; set; } = string.Empty;
        /// <summary>Original Excel defined name used to resolve source hyperlinks after target-name qualification.</summary>
        public string SourceName { get; set; } = string.Empty;
        public string? SheetName { get; set; }
        public string A1Range { get; set; } = string.Empty;
    }

    /// <summary>
    /// Adds a protected sheet/range to the target spreadsheet.
    /// </summary>
    public sealed class GoogleSheetsAddProtectedRangeRequest : GoogleSheetsRequest {
        public GoogleSheetsAddProtectedRangeRequest() : base("addProtectedRange") {
        }

        public string SheetName { get; set; } = string.Empty;
        public string? Description { get; set; }
        public bool WarningOnly { get; set; }
        public bool DomainUsersCanEdit { get; set; }
        public IReadOnlyList<string> EditorEmailAddresses { get; set; } = Array.Empty<string>();
        public IReadOnlyList<string> UnprotectedA1Ranges { get; set; } = Array.Empty<string>();
    }

    /// <summary>
    /// Applies a Google Sheets basic filter to a range.
    /// </summary>
    public sealed class GoogleSheetsSetBasicFilterRequest : GoogleSheetsRequest {
        public GoogleSheetsSetBasicFilterRequest() : base("setBasicFilter") {
        }

        public string SheetName { get; set; } = string.Empty;
        public string A1Range { get; set; } = string.Empty;
        public int StartRowIndex { get; set; }
        public int EndRowIndexExclusive { get; set; }
        public int StartColumnIndex { get; set; }
        public int EndColumnIndexExclusive { get; set; }
        public IReadOnlyList<GoogleSheetsFilterColumnCriteria> Criteria { get; set; } = Array.Empty<GoogleSheetsFilterColumnCriteria>();
    }

    /// <summary>
    /// Adds a filter view to a worksheet range.
    /// </summary>
    public sealed class GoogleSheetsAddFilterViewRequest : GoogleSheetsRequest {
        public GoogleSheetsAddFilterViewRequest() : base("addFilterView") {
        }

        public string SheetName { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string A1Range { get; set; } = string.Empty;
        public int StartRowIndex { get; set; }
        public int EndRowIndexExclusive { get; set; }
        public int StartColumnIndex { get; set; }
        public int EndColumnIndexExclusive { get; set; }
        public IReadOnlyList<GoogleSheetsFilterColumnCriteria> Criteria { get; set; } = Array.Empty<GoogleSheetsFilterColumnCriteria>();
    }

    /// <summary>
    /// Adds a native Google Sheets table to the worksheet.
    /// </summary>
    public sealed class GoogleSheetsAddTableRequest : GoogleSheetsRequest {
        public GoogleSheetsAddTableRequest() : base("addTable") {
        }

        public string SheetName { get; set; } = string.Empty;
        public string TableName { get; set; } = string.Empty;
        public string A1Range { get; set; } = string.Empty;
        public int StartRowIndex { get; set; }
        public int EndRowIndexExclusive { get; set; }
        public int StartColumnIndex { get; set; }
        public int EndColumnIndexExclusive { get; set; }
        public bool HasHeaderRow { get; set; }
        public bool TotalsRowShown { get; set; }
        public string? StyleName { get; set; }
        public string? HeaderColorArgb { get; set; }
        public string? FirstBandColorArgb { get; set; }
        public string? SecondBandColorArgb { get; set; }
        public string? FooterColorArgb { get; set; }
        public IReadOnlyList<GoogleSheetsTableColumn> Columns { get; set; } = Array.Empty<GoogleSheetsTableColumn>();
    }

    /// <summary>
    /// Merges a rectangular range in the target worksheet.
    /// </summary>
    public sealed class GoogleSheetsMergeCellsRequest : GoogleSheetsRequest {
        public GoogleSheetsMergeCellsRequest() : base("mergeCells") {
        }

        public string SheetName { get; set; } = string.Empty;
        public string A1Range { get; set; } = string.Empty;
        public int StartRowIndex { get; set; }
        public int EndRowIndexExclusive { get; set; }
        public int StartColumnIndex { get; set; }
        public int EndColumnIndexExclusive { get; set; }
    }

    /// <summary>
    /// Updates row or column dimension properties in the target worksheet.
    /// </summary>
    public sealed class GoogleSheetsUpdateDimensionPropertiesRequest : GoogleSheetsRequest {
        public GoogleSheetsUpdateDimensionPropertiesRequest() : base("updateDimensionProperties") {
        }

        public string SheetName { get; set; } = string.Empty;
        public GoogleSheetsDimensionKind DimensionKind { get; set; }
        public int StartIndex { get; set; }
        public int EndIndexExclusive { get; set; }
        public int? PixelSize { get; set; }
        public bool Hidden { get; set; }
        public byte? OutlineLevel { get; set; }
    }

    /// <summary>Adds a row or column outline group.</summary>
    public sealed class GoogleSheetsAddDimensionGroupRequest : GoogleSheetsRequest {
        public GoogleSheetsAddDimensionGroupRequest() : base("addDimensionGroup") { }
        public string SheetName { get; set; } = string.Empty;
        public GoogleSheetsDimensionKind DimensionKind { get; set; }
        public int StartIndex { get; set; }
        public int EndIndexExclusive { get; set; }
    }

    /// <summary>Adds one supported conditional-formatting rule.</summary>
    public sealed class GoogleSheetsAddConditionalFormatRuleRequest : GoogleSheetsRequest {
        public GoogleSheetsAddConditionalFormatRuleRequest() : base("addConditionalFormatRule") { }
        public string SheetName { get; set; } = string.Empty;
        public string A1Range { get; set; } = string.Empty;
        public int Index { get; set; }
        public string ConditionType { get; set; } = string.Empty;
        public IReadOnlyList<string> Values { get; set; } = Array.Empty<string>();
        public GoogleSheetsCellStyle? Format { get; set; }
    }

    /// <summary>Adds a chart backed by a generated hidden data range.</summary>
    public sealed class GoogleSheetsAddChartRequest : GoogleSheetsRequest {
        public GoogleSheetsAddChartRequest() : base("addChart") { }
        public string SheetName { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string ChartType { get; set; } = string.Empty;
        public string DataSheetName { get; set; } = string.Empty;
        public int DataStartRowIndex { get; set; }
        public int DataRowCount { get; set; }
        public int SeriesCount { get; set; }
        public int AnchorRowIndex { get; set; }
        public int AnchorColumnIndex { get; set; }
    }

    /// <summary>Adds a supported pivot table at a destination cell.</summary>
    public sealed class GoogleSheetsAddPivotTableRequest : GoogleSheetsRequest {
        public GoogleSheetsAddPivotTableRequest() : base("addPivotTable") { }
        public string SheetName { get; set; } = string.Empty;
        public int DestinationRowIndex { get; set; }
        public int DestinationColumnIndex { get; set; }
        public string SourceSheetName { get; set; } = string.Empty;
        public string SourceA1Range { get; set; } = string.Empty;
        public IReadOnlyList<GoogleSheetsPivotGroup> Rows { get; set; } = Array.Empty<GoogleSheetsPivotGroup>();
        public IReadOnlyList<GoogleSheetsPivotGroup> Columns { get; set; } = Array.Empty<GoogleSheetsPivotGroup>();
        public IReadOnlyList<GoogleSheetsPivotValue> Values { get; set; } = Array.Empty<GoogleSheetsPivotValue>();
    }

    public sealed class GoogleSheetsPivotGroup {
        public int SourceColumnOffset { get; set; }
        public bool ShowTotals { get; set; } = true;
        public string SortOrder { get; set; } = "ASCENDING";
    }

    public sealed class GoogleSheetsPivotValue {
        public int SourceColumnOffset { get; set; }
        public string SummarizeFunction { get; set; } = "SUM";
        public string? Name { get; set; }
    }

    /// <summary>Deletes spreadsheet-scoped developer metadata matching a key.</summary>
    public sealed class GoogleSheetsDeleteDeveloperMetadataRequest : GoogleSheetsRequest {
        public GoogleSheetsDeleteDeveloperMetadataRequest() : base("deleteDeveloperMetadata") { }
        public string Key { get; set; } = string.Empty;
    }

    /// <summary>Creates spreadsheet-scoped developer metadata.</summary>
    public sealed class GoogleSheetsCreateDeveloperMetadataRequest : GoogleSheetsRequest {
        public GoogleSheetsCreateDeveloperMetadataRequest() : base("createDeveloperMetadata") { }
        public string Key { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }

    /// <summary>
    /// Dimension kind used by dimension property requests.
    /// </summary>
    public enum GoogleSheetsDimensionKind {
        Rows = 0,
        Columns = 1,
    }

    /// <summary>
    /// Cell payload for the provider-neutral Google Sheets batch.
    /// Row and column indexes are zero-based to align with the Google Sheets API.
    /// </summary>
    public sealed class GoogleSheetsCellData {
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public GoogleSheetsCellValue Value { get; set; } = GoogleSheetsCellValue.Blank();
        public string? NumberFormatHint { get; set; }
        public GoogleSheetsCellStyle? Style { get; set; }
        public GoogleSheetsDataValidationRule? DataValidationRule { get; set; }
        public GoogleSheetsHyperlink? Hyperlink { get; set; }
        public GoogleSheetsComment? Comment { get; set; }
        public IReadOnlyList<GoogleSheetsTextFormatRun> TextFormatRuns { get; set; } = Array.Empty<GoogleSheetsTextFormatRun>();
    }

    /// <summary>
    /// Normalized cell value kinds used by the compiler output.
    /// </summary>
    public enum GoogleSheetsCellValueKind {
        Blank = 0,
        String = 1,
        Number = 2,
        Boolean = 3,
        DateTime = 4,
        Formula = 5,
    }

    /// <summary>
    /// Normalized value payload used in compiled Google Sheets requests.
    /// </summary>
    public sealed class GoogleSheetsCellValue {
        private GoogleSheetsCellValue(GoogleSheetsCellValueKind kind, object? value) {
            Kind = kind;
            Value = value;
        }

        public GoogleSheetsCellValueKind Kind { get; }
        public object? Value { get; }

        public static GoogleSheetsCellValue Blank() => new GoogleSheetsCellValue(GoogleSheetsCellValueKind.Blank, null);
        public static GoogleSheetsCellValue String(string? value) => new GoogleSheetsCellValue(GoogleSheetsCellValueKind.String, value ?? string.Empty);
        public static GoogleSheetsCellValue Number(double value) => new GoogleSheetsCellValue(GoogleSheetsCellValueKind.Number, value);
        public static GoogleSheetsCellValue Boolean(bool value) => new GoogleSheetsCellValue(GoogleSheetsCellValueKind.Boolean, value);
        public static GoogleSheetsCellValue DateTime(DateTime value) => new GoogleSheetsCellValue(GoogleSheetsCellValueKind.DateTime, value);
        public static GoogleSheetsCellValue Formula(string formula) => new GoogleSheetsCellValue(GoogleSheetsCellValueKind.Formula, formula ?? string.Empty);
    }

    /// <summary>
    /// Normalized cell formatting payload used by the compiler output.
    /// </summary>
    public sealed class GoogleSheetsCellStyle {
        public uint SourceStyleIndex { get; set; }
        public uint NumberFormatId { get; set; }
        public string? NumberFormatCode { get; set; }
        public bool IsDateLike { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikethrough { get; set; }
        public string? FontName { get; set; }
        public double? FontSize { get; set; }
        public string? FontColorArgb { get; set; }
        public string? FillColorArgb { get; set; }
        public GoogleSheetsCellBorders? Borders { get; set; }
        public string? HorizontalAlignment { get; set; }
        public string? VerticalAlignment { get; set; }
        public bool WrapText { get; set; }
        public int? TextRotation { get; set; }
        public uint? TextIndent { get; set; }
    }

    /// <summary>Rich-text run formatting beginning at a UTF-16 character index.</summary>
    public sealed class GoogleSheetsTextFormatRun {
        public int StartIndex { get; set; }
        public GoogleSheetsCellStyle Format { get; set; } = new GoogleSheetsCellStyle();
    }

    /// <summary>
    /// Provider-neutral border payload used by the Google Sheets compiler output.
    /// </summary>
    public sealed class GoogleSheetsCellBorders {
        public GoogleSheetsBorderSide? Left { get; set; }
        public GoogleSheetsBorderSide? Right { get; set; }
        public GoogleSheetsBorderSide? Top { get; set; }
        public GoogleSheetsBorderSide? Bottom { get; set; }
    }

    /// <summary>
    /// Provider-neutral single-border-side payload.
    /// </summary>
    public sealed class GoogleSheetsBorderSide {
        public string? Style { get; set; }
        public string? ColorArgb { get; set; }
    }

    /// <summary>
    /// Provider-neutral comment payload used by the Google Sheets compiler output.
    /// </summary>
    public sealed class GoogleSheetsComment {
        public string? Author { get; set; }
        public string Text { get; set; } = string.Empty;
    }

    /// <summary>
    /// Filter criteria keyed by the absolute zero-based sheet column index required by the Google Sheets API.
    /// </summary>
    public sealed class GoogleSheetsFilterColumnCriteria {
        public int ColumnId { get; set; }
        public IReadOnlyList<string> HiddenValues { get; set; } = Array.Empty<string>();
        public GoogleSheetsBooleanCondition? Condition { get; set; }
    }

    /// <summary>
    /// Provider-neutral boolean condition used by Google Sheets filter criteria.
    /// </summary>
    public sealed class GoogleSheetsBooleanCondition {
        public string Type { get; set; } = string.Empty;
        public IReadOnlyList<string> Values { get; set; } = Array.Empty<string>();
    }

    /// <summary>
    /// Provider-neutral table-column metadata for Google Sheets tables.
    /// </summary>
    public sealed class GoogleSheetsTableColumn {
        public int ColumnIndex { get; set; }
        public string Name { get; set; } = string.Empty;
        public string? ColumnType { get; set; }
        public string? TotalsRowFunction { get; set; }
        public GoogleSheetsDataValidationRule? DataValidationRule { get; set; }
    }

    /// <summary>
    /// Provider-neutral table-column validation metadata for native Google Sheets tables.
    /// </summary>
    public sealed class GoogleSheetsDataValidationRule {
        public string ConditionType { get; set; } = string.Empty;
        public IReadOnlyList<string> Values { get; set; } = Array.Empty<string>();
        public bool Strict { get; set; }
        public bool ShowCustomUi { get; set; }
    }

    /// <summary>
    /// Hyperlink metadata attached to a compiled cell payload.
    /// </summary>
    public sealed class GoogleSheetsHyperlink {
        public bool IsExternal { get; set; }
        public string Target { get; set; } = string.Empty;
    }
}
