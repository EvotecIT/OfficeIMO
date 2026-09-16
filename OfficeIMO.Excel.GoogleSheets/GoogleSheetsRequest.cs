namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>Updates target spreadsheet locale, time zone, and calculation settings.</summary>
    public sealed class GoogleSheetsUpdateSpreadsheetPropertiesRequest : GoogleSheetsRequest {
        /// <summary>Creates a spreadsheet-properties request.</summary>
        public GoogleSheetsUpdateSpreadsheetPropertiesRequest() : base("updateSpreadsheetProperties") { }
        /// <summary>Gets or sets the target locale.</summary>
        public string? Locale { get; set; }
        /// <summary>Gets or sets the target time zone.</summary>
        public string? TimeZone { get; set; }
        /// <summary>Gets or sets the requested recalculation interval.</summary>
        public GoogleSheetsRecalculationInterval RecalculationInterval { get; set; }
    }

    /// <summary>
    /// Base request emitted by the Google Sheets batch compiler.
    /// </summary>
    public abstract class GoogleSheetsRequest {
        /// <summary>Initializes the provider-neutral request kind.</summary>
        protected GoogleSheetsRequest(string kind) {
            Kind = kind ?? throw new ArgumentNullException(nameof(kind));
        }

        /// <summary>Gets the batch request kind used by the payload builder.</summary>
        public string Kind { get; }
    }

    /// <summary>
    /// Adds or configures a worksheet in the target spreadsheet.
    /// </summary>
    public sealed class GoogleSheetsAddSheetRequest : GoogleSheetsRequest {
        /// <summary>Creates an add-sheet request.</summary>
        public GoogleSheetsAddSheetRequest() : base("addSheet") {
        }

        /// <summary>Gets or sets the worksheet name.</summary>
        public string SheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets the zero-based sheet position.</summary>
        public int SheetIndex { get; set; }
        /// <summary>Gets or sets whether the sheet is hidden.</summary>
        public bool Hidden { get; set; }
        /// <summary>Gets or sets whether the sheet reads right to left.</summary>
        public bool RightToLeft { get; set; }
        /// <summary>Gets or sets an optional ARGB tab color.</summary>
        public string? TabColorArgb { get; set; }
        /// <summary>Gets or sets the sheet grid's row count.</summary>
        public int RowCount { get; set; }
        /// <summary>Gets or sets the sheet grid's column count.</summary>
        public int ColumnCount { get; set; }
        /// <summary>Gets or sets the number of frozen top rows.</summary>
        public int FrozenRowCount { get; set; }
        /// <summary>Gets or sets the number of frozen left columns.</summary>
        public int FrozenColumnCount { get; set; }
        /// <summary>Gets or sets whether gridlines are hidden.</summary>
        public bool HideGridlines { get; set; }
    }

    /// <summary>
    /// Writes cell values into a worksheet.
    /// </summary>
    public sealed class GoogleSheetsUpdateCellsRequest : GoogleSheetsRequest {
        private readonly List<GoogleSheetsCellData> _cells = new List<GoogleSheetsCellData>();

        /// <summary>Creates an update-cells request.</summary>
        public GoogleSheetsUpdateCellsRequest() : base("updateCells") {
        }

        /// <summary>Gets or sets the worksheet receiving the cell data.</summary>
        public string SheetName { get; set; } = string.Empty;
        /// <summary>Gets the cells accumulated by the batch compiler.</summary>
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
        /// <summary>Creates a range validation request.</summary>
        public GoogleSheetsSetDataValidationRequest() : base("setDataValidation") {
        }

        /// <summary>Gets or sets the target worksheet.</summary>
        public string SheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets the range in A1 notation.</summary>
        public string A1Range { get; set; } = string.Empty;
        /// <summary>Gets or sets the inclusive zero-based first row.</summary>
        public int StartRowIndex { get; set; }
        /// <summary>Gets or sets the exclusive zero-based ending row.</summary>
        public int EndRowIndexExclusive { get; set; }
        /// <summary>Gets or sets the inclusive zero-based first column.</summary>
        public int StartColumnIndex { get; set; }
        /// <summary>Gets or sets the exclusive zero-based ending column.</summary>
        public int EndColumnIndexExclusive { get; set; }
        /// <summary>Gets or sets the validation rule applied across the range.</summary>
        public GoogleSheetsDataValidationRule Rule { get; set; } = new GoogleSheetsDataValidationRule();
    }

    /// <summary>
    /// Adds a named range to the target spreadsheet.
    /// </summary>
    public sealed class GoogleSheetsAddNamedRangeRequest : GoogleSheetsRequest {
        /// <summary>Creates an add-named-range request.</summary>
        public GoogleSheetsAddNamedRangeRequest() : base("addNamedRange") {
        }

        /// <summary>Gets or sets the target named-range name.</summary>
        public string Name { get; set; } = string.Empty;
        /// <summary>Original Excel defined name used to resolve source hyperlinks after target-name qualification.</summary>
        public string SourceName { get; set; } = string.Empty;
        /// <summary>Gets or sets the owning sheet, when the name is sheet-scoped.</summary>
        public string? SheetName { get; set; }
        /// <summary>Gets or sets the referenced range in A1 notation.</summary>
        public string A1Range { get; set; } = string.Empty;
    }

    /// <summary>
    /// Adds a protected sheet/range to the target spreadsheet.
    /// </summary>
    public sealed class GoogleSheetsAddProtectedRangeRequest : GoogleSheetsRequest {
        /// <summary>Creates a protected-range request.</summary>
        public GoogleSheetsAddProtectedRangeRequest() : base("addProtectedRange") {
        }

        /// <summary>Gets or sets the protected worksheet.</summary>
        public string SheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets an optional range description.</summary>
        public string? Description { get; set; }
        /// <summary>Gets or sets whether edits warn instead of being blocked.</summary>
        public bool WarningOnly { get; set; }
        /// <summary>Gets or sets whether domain users may edit the range.</summary>
        public bool DomainUsersCanEdit { get; set; }
        /// <summary>Gets or sets the permitted editor email addresses.</summary>
        public IReadOnlyList<string> EditorEmailAddresses { get; set; } = Array.Empty<string>();
        /// <summary>Gets or sets unprotected A1 ranges within the protected sheet.</summary>
        public IReadOnlyList<string> UnprotectedA1Ranges { get; set; } = Array.Empty<string>();
    }

    /// <summary>
    /// Applies a Google Sheets basic filter to a range.
    /// </summary>
    public sealed class GoogleSheetsSetBasicFilterRequest : GoogleSheetsRequest {
        /// <summary>Creates a basic-filter request.</summary>
        public GoogleSheetsSetBasicFilterRequest() : base("setBasicFilter") {
        }

        /// <summary>Gets or sets the worksheet containing the filter.</summary>
        public string SheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets the filter range in A1 notation.</summary>
        public string A1Range { get; set; } = string.Empty;
        /// <summary>Gets or sets the inclusive zero-based first row.</summary>
        public int StartRowIndex { get; set; }
        /// <summary>Gets or sets the exclusive zero-based ending row.</summary>
        public int EndRowIndexExclusive { get; set; }
        /// <summary>Gets or sets the inclusive zero-based first column.</summary>
        public int StartColumnIndex { get; set; }
        /// <summary>Gets or sets the exclusive zero-based ending column.</summary>
        public int EndColumnIndexExclusive { get; set; }
        /// <summary>Gets or sets the column-specific filter criteria.</summary>
        public IReadOnlyList<GoogleSheetsFilterColumnCriteria> Criteria { get; set; } = Array.Empty<GoogleSheetsFilterColumnCriteria>();
    }

    /// <summary>
    /// Adds a filter view to a worksheet range.
    /// </summary>
    public sealed class GoogleSheetsAddFilterViewRequest : GoogleSheetsRequest {
        /// <summary>Creates a filter-view request.</summary>
        public GoogleSheetsAddFilterViewRequest() : base("addFilterView") {
        }

        /// <summary>Gets or sets the worksheet containing the view.</summary>
        public string SheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets the filter-view title.</summary>
        public string Title { get; set; } = string.Empty;
        /// <summary>Gets or sets the view range in A1 notation.</summary>
        public string A1Range { get; set; } = string.Empty;
        /// <summary>Gets or sets the inclusive zero-based first row.</summary>
        public int StartRowIndex { get; set; }
        /// <summary>Gets or sets the exclusive zero-based ending row.</summary>
        public int EndRowIndexExclusive { get; set; }
        /// <summary>Gets or sets the inclusive zero-based first column.</summary>
        public int StartColumnIndex { get; set; }
        /// <summary>Gets or sets the exclusive zero-based ending column.</summary>
        public int EndColumnIndexExclusive { get; set; }
        /// <summary>Gets or sets the column-specific filter criteria.</summary>
        public IReadOnlyList<GoogleSheetsFilterColumnCriteria> Criteria { get; set; } = Array.Empty<GoogleSheetsFilterColumnCriteria>();
    }

    /// <summary>
    /// Adds a native Google Sheets table to the worksheet.
    /// </summary>
    public sealed class GoogleSheetsAddTableRequest : GoogleSheetsRequest {
        /// <summary>Creates an add-table request.</summary>
        public GoogleSheetsAddTableRequest() : base("addTable") {
        }

        /// <summary>Gets or sets the worksheet containing the table.</summary>
        public string SheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets the native table name.</summary>
        public string TableName { get; set; } = string.Empty;
        /// <summary>Gets or sets the table range in A1 notation.</summary>
        public string A1Range { get; set; } = string.Empty;
        /// <summary>Gets or sets the inclusive zero-based first row.</summary>
        public int StartRowIndex { get; set; }
        /// <summary>Gets or sets the exclusive zero-based ending row.</summary>
        public int EndRowIndexExclusive { get; set; }
        /// <summary>Gets or sets the inclusive zero-based first column.</summary>
        public int StartColumnIndex { get; set; }
        /// <summary>Gets or sets the exclusive zero-based ending column.</summary>
        public int EndColumnIndexExclusive { get; set; }
        /// <summary>Gets or sets whether the first row contains headers.</summary>
        public bool HasHeaderRow { get; set; }
        /// <summary>Gets or sets whether the totals row is shown.</summary>
        public bool TotalsRowShown { get; set; }
        /// <summary>Gets or sets the source table-style name, when available.</summary>
        public string? StyleName { get; set; }
        /// <summary>Gets or sets the optional ARGB header color.</summary>
        public string? HeaderColorArgb { get; set; }
        /// <summary>Gets or sets the optional ARGB first-band color.</summary>
        public string? FirstBandColorArgb { get; set; }
        /// <summary>Gets or sets the optional ARGB second-band color.</summary>
        public string? SecondBandColorArgb { get; set; }
        /// <summary>Gets or sets the optional ARGB footer color.</summary>
        public string? FooterColorArgb { get; set; }
        /// <summary>Gets or sets target table-column metadata.</summary>
        public IReadOnlyList<GoogleSheetsTableColumn> Columns { get; set; } = Array.Empty<GoogleSheetsTableColumn>();
    }

    /// <summary>
    /// Merges a rectangular range in the target worksheet.
    /// </summary>
    public sealed class GoogleSheetsMergeCellsRequest : GoogleSheetsRequest {
        /// <summary>Creates a merge-cells request.</summary>
        public GoogleSheetsMergeCellsRequest() : base("mergeCells") {
        }

        /// <summary>Gets or sets the worksheet containing the merge.</summary>
        public string SheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets the merged range in A1 notation.</summary>
        public string A1Range { get; set; } = string.Empty;
        /// <summary>Gets or sets the inclusive zero-based first row.</summary>
        public int StartRowIndex { get; set; }
        /// <summary>Gets or sets the exclusive zero-based ending row.</summary>
        public int EndRowIndexExclusive { get; set; }
        /// <summary>Gets or sets the inclusive zero-based first column.</summary>
        public int StartColumnIndex { get; set; }
        /// <summary>Gets or sets the exclusive zero-based ending column.</summary>
        public int EndColumnIndexExclusive { get; set; }
    }

    /// <summary>
    /// Updates row or column dimension properties in the target worksheet.
    /// </summary>
    public sealed class GoogleSheetsUpdateDimensionPropertiesRequest : GoogleSheetsRequest {
        /// <summary>Creates a row- or column-dimension update request.</summary>
        public GoogleSheetsUpdateDimensionPropertiesRequest() : base("updateDimensionProperties") {
        }

        /// <summary>Gets or sets the worksheet containing the dimensions.</summary>
        public string SheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets whether the range covers rows or columns.</summary>
        public GoogleSheetsDimensionKind DimensionKind { get; set; }
        /// <summary>Gets or sets the inclusive zero-based first dimension index.</summary>
        public int StartIndex { get; set; }
        /// <summary>Gets or sets the exclusive zero-based ending dimension index.</summary>
        public int EndIndexExclusive { get; set; }
        /// <summary>Gets or sets the optional dimension size in pixels.</summary>
        public int? PixelSize { get; set; }
        /// <summary>Gets or sets whether the dimensions are hidden.</summary>
        public bool Hidden { get; set; }
        /// <summary>Gets or sets the source outline level, when available.</summary>
        public byte? OutlineLevel { get; set; }
    }

    /// <summary>Adds a row or column outline group.</summary>
    public sealed class GoogleSheetsAddDimensionGroupRequest : GoogleSheetsRequest {
        /// <summary>Creates an outline-group request.</summary>
        public GoogleSheetsAddDimensionGroupRequest() : base("addDimensionGroup") { }
        /// <summary>Gets or sets the worksheet containing the group.</summary>
        public string SheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets whether the group covers rows or columns.</summary>
        public GoogleSheetsDimensionKind DimensionKind { get; set; }
        /// <summary>Gets or sets the inclusive zero-based first dimension index.</summary>
        public int StartIndex { get; set; }
        /// <summary>Gets or sets the exclusive zero-based ending dimension index.</summary>
        public int EndIndexExclusive { get; set; }
    }

    /// <summary>Adds one supported conditional-formatting rule.</summary>
    public sealed class GoogleSheetsAddConditionalFormatRuleRequest : GoogleSheetsRequest {
        /// <summary>Creates a conditional-formatting request.</summary>
        public GoogleSheetsAddConditionalFormatRuleRequest() : base("addConditionalFormatRule") { }
        /// <summary>Gets or sets the worksheet containing the rule.</summary>
        public string SheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets the formatted range in A1 notation.</summary>
        public string A1Range { get; set; } = string.Empty;
        /// <summary>Gets or sets the rule insertion index.</summary>
        public int Index { get; set; }
        /// <summary>Gets or sets the Sheets boolean-condition type.</summary>
        public string ConditionType { get; set; } = string.Empty;
        /// <summary>Gets or sets the condition's argument values.</summary>
        public IReadOnlyList<string> Values { get; set; } = Array.Empty<string>();
        /// <summary>Gets or sets the formatting applied when the condition matches.</summary>
        public GoogleSheetsCellStyle? Format { get; set; }
    }

    /// <summary>Adds a chart backed by a generated hidden data range.</summary>
    public sealed class GoogleSheetsAddChartRequest : GoogleSheetsRequest {
        /// <summary>Creates a chart request backed by hidden sheet data.</summary>
        public GoogleSheetsAddChartRequest() : base("addChart") { }
        /// <summary>Gets or sets the worksheet receiving the chart.</summary>
        public string SheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets the chart title.</summary>
        public string Title { get; set; } = string.Empty;
        /// <summary>Gets or sets the mapped Sheets chart family.</summary>
        public string ChartType { get; set; } = string.Empty;
        /// <summary>Gets or sets the generated data worksheet name.</summary>
        public string DataSheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets the zero-based first data row.</summary>
        public int DataStartRowIndex { get; set; }
        /// <summary>Gets or sets the number of data rows.</summary>
        public int DataRowCount { get; set; }
        /// <summary>Gets or sets the number of chart series.</summary>
        public int SeriesCount { get; set; }
        /// <summary>Gets or sets the zero-based chart-anchor row.</summary>
        public int AnchorRowIndex { get; set; }
        /// <summary>Gets or sets the zero-based chart-anchor column.</summary>
        public int AnchorColumnIndex { get; set; }
    }

    /// <summary>Adds a supported pivot table at a destination cell.</summary>
    public sealed class GoogleSheetsAddPivotTableRequest : GoogleSheetsRequest {
        /// <summary>Creates a pivot-table request.</summary>
        public GoogleSheetsAddPivotTableRequest() : base("addPivotTable") { }
        /// <summary>Gets or sets the worksheet receiving the pivot table.</summary>
        public string SheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets the zero-based destination row.</summary>
        public int DestinationRowIndex { get; set; }
        /// <summary>Gets or sets the zero-based destination column.</summary>
        public int DestinationColumnIndex { get; set; }
        /// <summary>Gets or sets the source-data worksheet.</summary>
        public string SourceSheetName { get; set; } = string.Empty;
        /// <summary>Gets or sets the source-data range in A1 notation.</summary>
        public string SourceA1Range { get; set; } = string.Empty;
        /// <summary>Gets or sets source columns grouped as pivot rows.</summary>
        public IReadOnlyList<GoogleSheetsPivotGroup> Rows { get; set; } = Array.Empty<GoogleSheetsPivotGroup>();
        /// <summary>Gets or sets source columns grouped as pivot columns.</summary>
        public IReadOnlyList<GoogleSheetsPivotGroup> Columns { get; set; } = Array.Empty<GoogleSheetsPivotGroup>();
        /// <summary>Gets or sets aggregated pivot values.</summary>
        public IReadOnlyList<GoogleSheetsPivotValue> Values { get; set; } = Array.Empty<GoogleSheetsPivotValue>();
    }

    /// <summary>A source column grouped on a pivot-table axis.</summary>
    public sealed class GoogleSheetsPivotGroup {
        /// <summary>Gets or sets the column offset within the source range.</summary>
        public int SourceColumnOffset { get; set; }
        /// <summary>Gets or sets whether totals are shown; defaults to true.</summary>
        public bool ShowTotals { get; set; } = true;
        /// <summary>Gets or sets the Sheets sort order; defaults to <c>ASCENDING</c>.</summary>
        public string SortOrder { get; set; } = "ASCENDING";
    }

    /// <summary>A source column aggregated in a pivot table.</summary>
    public sealed class GoogleSheetsPivotValue {
        /// <summary>Gets or sets the column offset within the source range.</summary>
        public int SourceColumnOffset { get; set; }
        /// <summary>Gets or sets the Sheets summary function; defaults to <c>SUM</c>.</summary>
        public string SummarizeFunction { get; set; } = "SUM";
        /// <summary>Gets or sets an optional display name for the value.</summary>
        public string? Name { get; set; }
    }

    /// <summary>Deletes spreadsheet-scoped developer metadata matching a key.</summary>
    public sealed class GoogleSheetsDeleteDeveloperMetadataRequest : GoogleSheetsRequest {
        /// <summary>Creates a metadata-deletion request.</summary>
        public GoogleSheetsDeleteDeveloperMetadataRequest() : base("deleteDeveloperMetadata") { }
        /// <summary>Gets or sets the spreadsheet-scoped metadata key to delete.</summary>
        public string Key { get; set; } = string.Empty;
    }

    /// <summary>Creates spreadsheet-scoped developer metadata.</summary>
    public sealed class GoogleSheetsCreateDeveloperMetadataRequest : GoogleSheetsRequest {
        /// <summary>Creates a metadata-write request.</summary>
        public GoogleSheetsCreateDeveloperMetadataRequest() : base("createDeveloperMetadata") { }
        /// <summary>Gets or sets the spreadsheet-scoped metadata key.</summary>
        public string Key { get; set; } = string.Empty;
        /// <summary>Gets or sets the metadata value.</summary>
        public string Value { get; set; } = string.Empty;
    }

    /// <summary>
    /// Dimension kind used by dimension property requests.
    /// </summary>
    public enum GoogleSheetsDimensionKind {
        /// <summary>Worksheet rows.</summary>
        Rows = 0,
        /// <summary>Worksheet columns.</summary>
        Columns = 1,
    }

    /// <summary>
    /// Cell payload for the provider-neutral Google Sheets batch.
    /// Row and column indexes are zero-based to align with the Google Sheets API.
    /// </summary>
    public sealed class GoogleSheetsCellData {
        /// <summary>Gets or sets the zero-based row index.</summary>
        public int RowIndex { get; set; }
        /// <summary>Gets or sets the zero-based column index.</summary>
        public int ColumnIndex { get; set; }
        /// <summary>Gets or sets the normalized cell value; defaults to blank.</summary>
        public GoogleSheetsCellValue Value { get; set; } = GoogleSheetsCellValue.Blank();
        /// <summary>Gets or sets an optional number-format hint from the source cell.</summary>
        public string? NumberFormatHint { get; set; }
        /// <summary>Gets or sets translated cell formatting.</summary>
        public GoogleSheetsCellStyle? Style { get; set; }
        /// <summary>Gets or sets a validation rule associated with this cell.</summary>
        public GoogleSheetsDataValidationRule? DataValidationRule { get; set; }
        /// <summary>Gets or sets hyperlink metadata associated with this cell.</summary>
        public GoogleSheetsHyperlink? Hyperlink { get; set; }
        /// <summary>Gets or sets a source comment flattened into a Sheets note.</summary>
        public GoogleSheetsComment? Comment { get; set; }
        /// <summary>Gets or sets rich-text runs indexed by UTF-16 offsets.</summary>
        public IReadOnlyList<GoogleSheetsTextFormatRun> TextFormatRuns { get; set; } = Array.Empty<GoogleSheetsTextFormatRun>();
    }

    /// <summary>
    /// Normalized cell value kinds used by the compiler output.
    /// </summary>
    public enum GoogleSheetsCellValueKind {
        /// <summary>No cell value.</summary>
        Blank = 0,
        /// <summary>A text value.</summary>
        String = 1,
        /// <summary>A numeric value.</summary>
        Number = 2,
        /// <summary>A Boolean value.</summary>
        Boolean = 3,
        /// <summary>A date/time value awaiting Sheets serial conversion.</summary>
        DateTime = 4,
        /// <summary>A formula string.</summary>
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

        /// <summary>Gets the normalized value category.</summary>
        public GoogleSheetsCellValueKind Kind { get; }
        /// <summary>Gets the boxed value, or null for a blank cell.</summary>
        public object? Value { get; }

        /// <summary>Creates a blank cell value.</summary>
        public static GoogleSheetsCellValue Blank() => new GoogleSheetsCellValue(GoogleSheetsCellValueKind.Blank, null);
        /// <summary>Creates text, treating null as empty.</summary>
        public static GoogleSheetsCellValue String(string? value) => new GoogleSheetsCellValue(GoogleSheetsCellValueKind.String, value ?? string.Empty);
        /// <summary>Creates a numeric cell value.</summary>
        public static GoogleSheetsCellValue Number(double value) => new GoogleSheetsCellValue(GoogleSheetsCellValueKind.Number, value);
        /// <summary>Creates a Boolean cell value.</summary>
        public static GoogleSheetsCellValue Boolean(bool value) => new GoogleSheetsCellValue(GoogleSheetsCellValueKind.Boolean, value);
        /// <summary>Creates a date/time cell value.</summary>
        public static GoogleSheetsCellValue DateTime(DateTime value) => new GoogleSheetsCellValue(GoogleSheetsCellValueKind.DateTime, value);
        /// <summary>Creates a formula cell value, treating null as empty.</summary>
        public static GoogleSheetsCellValue Formula(string formula) => new GoogleSheetsCellValue(GoogleSheetsCellValueKind.Formula, formula ?? string.Empty);
    }

    /// <summary>
    /// Normalized cell formatting payload used by the compiler output.
    /// </summary>
    public sealed class GoogleSheetsCellStyle {
        /// <summary>Gets or sets the source workbook style-table index.</summary>
        public uint SourceStyleIndex { get; set; }
        /// <summary>Gets or sets the source number-format identifier.</summary>
        public uint NumberFormatId { get; set; }
        /// <summary>Gets or sets the source number-format code, when available.</summary>
        public string? NumberFormatCode { get; set; }
        /// <summary>Gets or sets whether the source number format is date-like.</summary>
        public bool IsDateLike { get; set; }
        /// <summary>Gets or sets bold formatting.</summary>
        public bool Bold { get; set; }
        /// <summary>Gets or sets italic formatting.</summary>
        public bool Italic { get; set; }
        /// <summary>Gets or sets underline formatting.</summary>
        public bool Underline { get; set; }
        /// <summary>Gets or sets strikethrough formatting.</summary>
        public bool Strikethrough { get; set; }
        /// <summary>Gets or sets the font family.</summary>
        public string? FontName { get; set; }
        /// <summary>Gets or sets the font size.</summary>
        public double? FontSize { get; set; }
        /// <summary>Gets or sets the font color as ARGB text.</summary>
        public string? FontColorArgb { get; set; }
        /// <summary>Gets or sets the fill color as ARGB text.</summary>
        public string? FillColorArgb { get; set; }
        /// <summary>Gets or sets optional cell-border sides.</summary>
        public GoogleSheetsCellBorders? Borders { get; set; }
        /// <summary>Gets or sets horizontal alignment.</summary>
        public string? HorizontalAlignment { get; set; }
        /// <summary>Gets or sets vertical alignment.</summary>
        public string? VerticalAlignment { get; set; }
        /// <summary>Gets or sets whether text wraps within the cell.</summary>
        public bool WrapText { get; set; }
        /// <summary>Gets or sets text rotation in degrees.</summary>
        public int? TextRotation { get; set; }
        /// <summary>Gets or sets source text indentation.</summary>
        public uint? TextIndent { get; set; }
    }

    /// <summary>Rich-text run formatting beginning at a UTF-16 character index.</summary>
    public sealed class GoogleSheetsTextFormatRun {
        /// <summary>Gets or sets the run's starting UTF-16 offset within cell text.</summary>
        public int StartIndex { get; set; }
        /// <summary>Gets or sets the formatting applied from this offset.</summary>
        public GoogleSheetsCellStyle Format { get; set; } = new GoogleSheetsCellStyle();
    }

    /// <summary>
    /// Provider-neutral border payload used by the Google Sheets compiler output.
    /// </summary>
    public sealed class GoogleSheetsCellBorders {
        /// <summary>Gets or sets the left border.</summary>
        public GoogleSheetsBorderSide? Left { get; set; }
        /// <summary>Gets or sets the right border.</summary>
        public GoogleSheetsBorderSide? Right { get; set; }
        /// <summary>Gets or sets the top border.</summary>
        public GoogleSheetsBorderSide? Top { get; set; }
        /// <summary>Gets or sets the bottom border.</summary>
        public GoogleSheetsBorderSide? Bottom { get; set; }
    }

    /// <summary>
    /// Provider-neutral single-border-side payload.
    /// </summary>
    public sealed class GoogleSheetsBorderSide {
        /// <summary>Gets or sets the mapped border style.</summary>
        public string? Style { get; set; }
        /// <summary>Gets or sets the border color as ARGB text.</summary>
        public string? ColorArgb { get; set; }
    }

    /// <summary>
    /// Provider-neutral comment payload used by the Google Sheets compiler output.
    /// </summary>
    public sealed class GoogleSheetsComment {
        /// <summary>Gets or sets the source comment author, when present.</summary>
        public string? Author { get; set; }
        /// <summary>Gets or sets comment text written as a Sheets note.</summary>
        public string Text { get; set; } = string.Empty;
    }

    /// <summary>
    /// Filter criteria keyed by the absolute zero-based sheet column index required by the Google Sheets API.
    /// </summary>
    public sealed class GoogleSheetsFilterColumnCriteria {
        /// <summary>Gets or sets the absolute zero-based worksheet column index.</summary>
        public int ColumnId { get; set; }
        /// <summary>Gets or sets values hidden by the filter.</summary>
        public IReadOnlyList<string> HiddenValues { get; set; } = Array.Empty<string>();
        /// <summary>Gets or sets an optional condition for this column.</summary>
        public GoogleSheetsBooleanCondition? Condition { get; set; }
    }

    /// <summary>
    /// Provider-neutral boolean condition used by Google Sheets filter criteria.
    /// </summary>
    public sealed class GoogleSheetsBooleanCondition {
        /// <summary>Gets or sets the Sheets condition type.</summary>
        public string Type { get; set; } = string.Empty;
        /// <summary>Gets or sets the condition's argument values.</summary>
        public IReadOnlyList<string> Values { get; set; } = Array.Empty<string>();
    }

    /// <summary>
    /// Provider-neutral table-column metadata for Google Sheets tables.
    /// </summary>
    public sealed class GoogleSheetsTableColumn {
        /// <summary>Gets or sets the zero-based column position within the table.</summary>
        public int ColumnIndex { get; set; }
        /// <summary>Gets or sets the table-column name.</summary>
        public string Name { get; set; } = string.Empty;
        /// <summary>Gets or sets an optional Sheets column type.</summary>
        public string? ColumnType { get; set; }
        /// <summary>Gets or sets an optional totals-row aggregation function.</summary>
        public string? TotalsRowFunction { get; set; }
        /// <summary>Gets or sets a column validation rule, when present.</summary>
        public GoogleSheetsDataValidationRule? DataValidationRule { get; set; }
    }

    /// <summary>
    /// Provider-neutral table-column validation metadata for native Google Sheets tables.
    /// </summary>
    public sealed class GoogleSheetsDataValidationRule {
        /// <summary>Gets or sets the Sheets validation-condition type.</summary>
        public string ConditionType { get; set; } = string.Empty;
        /// <summary>Gets or sets the condition's argument values.</summary>
        public IReadOnlyList<string> Values { get; set; } = Array.Empty<string>();
        /// <summary>Gets or sets whether invalid input is rejected.</summary>
        public bool Strict { get; set; }
        /// <summary>Gets or sets whether the validation UI is displayed.</summary>
        public bool ShowCustomUi { get; set; }
    }

    /// <summary>
    /// Hyperlink metadata attached to a compiled cell payload.
    /// </summary>
    public sealed class GoogleSheetsHyperlink {
        /// <summary>Gets or sets whether the target points outside the spreadsheet.</summary>
        public bool IsExternal { get; set; }
        /// <summary>Gets or sets the link target.</summary>
        public string Target { get; set; } = string.Empty;
    }
}
