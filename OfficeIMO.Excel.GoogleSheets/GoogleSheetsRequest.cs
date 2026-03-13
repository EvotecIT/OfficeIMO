namespace OfficeIMO.Excel.GoogleSheets {
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
        public int FrozenRowCount { get; set; }
        public int FrozenColumnCount { get; set; }
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
    /// Adds a named range to the target spreadsheet.
    /// </summary>
    public sealed class GoogleSheetsAddNamedRangeRequest : GoogleSheetsRequest {
        public GoogleSheetsAddNamedRangeRequest() : base("addNamedRange") {
        }

        public string Name { get; set; } = string.Empty;
        public string? SheetName { get; set; }
        public string A1Range { get; set; } = string.Empty;
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
        public GoogleSheetsHyperlink? Hyperlink { get; set; }
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
        public string? FontColorArgb { get; set; }
        public string? FillColorArgb { get; set; }
        public GoogleSheetsCellBorders? Borders { get; set; }
        public string? HorizontalAlignment { get; set; }
        public string? VerticalAlignment { get; set; }
        public bool WrapText { get; set; }
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
    /// Filter criteria for one relative column within a Google Sheets filter range.
    /// </summary>
    public sealed class GoogleSheetsFilterColumnCriteria {
        public int ColumnId { get; set; }
        public IReadOnlyList<string> HiddenValues { get; set; } = Array.Empty<string>();
    }

    /// <summary>
    /// Provider-neutral table-column metadata for Google Sheets tables.
    /// </summary>
    public sealed class GoogleSheetsTableColumn {
        public int ColumnIndex { get; set; }
        public string Name { get; set; } = string.Empty;
        public string? ColumnType { get; set; }
    }

    /// <summary>
    /// Hyperlink metadata attached to a compiled cell payload.
    /// </summary>
    public sealed class GoogleSheetsHyperlink {
        public bool IsExternal { get; set; }
        public string Target { get; set; } = string.Empty;
    }
}
