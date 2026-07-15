using System.Text.Json.Serialization;

namespace OfficeIMO.Excel.GoogleSheets {
    internal sealed class GoogleSheetsApiCreateSpreadsheetPayload {
        [JsonPropertyName("properties")]
        public GoogleSheetsApiSpreadsheetPropertiesPayload Properties { get; set; } = new GoogleSheetsApiSpreadsheetPropertiesPayload();

        [JsonPropertyName("sheets")]
        public List<GoogleSheetsApiSheetPayload> Sheets { get; } = new List<GoogleSheetsApiSheetPayload>();
    }

    internal sealed class GoogleSheetsApiSpreadsheetPropertiesPayload {
        [JsonPropertyName("title")]
        public string? Title { get; set; }

        [JsonPropertyName("locale")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Locale { get; set; }

        [JsonPropertyName("timeZone")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? TimeZone { get; set; }

        [JsonPropertyName("autoRecalc")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? AutoRecalc { get; set; }
    }

    internal sealed class GoogleSheetsApiSheetPayload {
        [JsonPropertyName("properties")]
        public GoogleSheetsApiSheetPropertiesPayload Properties { get; set; } = new GoogleSheetsApiSheetPropertiesPayload();
    }

    internal sealed class GoogleSheetsApiSheetPropertiesPayload {
        [JsonPropertyName("sheetId")]
        public int SheetId { get; set; }

        [JsonPropertyName("title")]
        public string Title { get; set; } = string.Empty;

        [JsonPropertyName("index")]
        public int Index { get; set; }

        [JsonPropertyName("hidden")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool Hidden { get; set; }

        [JsonPropertyName("rightToLeft")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? RightToLeft { get; set; }

        [JsonPropertyName("tabColor")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiColorPayload? TabColor { get; set; }

        [JsonPropertyName("gridProperties")]
        public GoogleSheetsApiGridPropertiesPayload GridProperties { get; set; } = new GoogleSheetsApiGridPropertiesPayload();
    }

    internal sealed class GoogleSheetsApiGridPropertiesPayload {
        [JsonPropertyName("frozenRowCount")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? FrozenRowCount { get; set; }

        [JsonPropertyName("frozenColumnCount")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? FrozenColumnCount { get; set; }

        [JsonPropertyName("hideGridlines")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool HideGridlines { get; set; }
    }

    internal sealed class GoogleSheetsApiBatchUpdatePayload {
        [JsonPropertyName("requests")]
        public List<GoogleSheetsApiRequestPayload> Requests { get; } = new List<GoogleSheetsApiRequestPayload>();

        [JsonPropertyName("includeSpreadsheetInResponse")]
        public bool IncludeSpreadsheetInResponse { get; set; }
    }

    internal sealed class GoogleSheetsApiRequestPayload {
        [JsonPropertyName("addSheet")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiAddSheetRequestPayload? AddSheet { get; set; }

        [JsonPropertyName("updateCells")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiUpdateCellsRequestPayload? UpdateCells { get; set; }

        [JsonPropertyName("setDataValidation")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiSetDataValidationRequestPayload? SetDataValidation { get; set; }

        [JsonPropertyName("mergeCells")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiMergeCellsRequestPayload? MergeCells { get; set; }

        [JsonPropertyName("addNamedRange")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiAddNamedRangeRequestPayload? AddNamedRange { get; set; }

        [JsonPropertyName("addProtectedRange")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiAddProtectedRangeRequestPayload? AddProtectedRange { get; set; }

        [JsonPropertyName("updateDimensionProperties")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiUpdateDimensionPropertiesRequestPayload? UpdateDimensionProperties { get; set; }

        [JsonPropertyName("deleteSheet")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiDeleteSheetRequestPayload? DeleteSheet { get; set; }

        [JsonPropertyName("updateSpreadsheetProperties")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiUpdateSpreadsheetPropertiesRequestPayload? UpdateSpreadsheetProperties { get; set; }

        [JsonPropertyName("setBasicFilter")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiSetBasicFilterRequestPayload? SetBasicFilter { get; set; }

        [JsonPropertyName("addFilterView")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiAddFilterViewRequestPayload? AddFilterView { get; set; }

        [JsonPropertyName("addTable")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiAddTableRequestPayload? AddTable { get; set; }

        [JsonPropertyName("addConditionalFormatRule")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiAddConditionalFormatRuleRequestPayload? AddConditionalFormatRule { get; set; }

        [JsonPropertyName("addChart")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiAddChartRequestPayload? AddChart { get; set; }

        [JsonPropertyName("addDimensionGroup")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiAddDimensionGroupRequestPayload? AddDimensionGroup { get; set; }

        [JsonPropertyName("addDeveloperMetadata")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiAddDeveloperMetadataRequestPayload? AddDeveloperMetadata { get; set; }
    }

    internal sealed class GoogleSheetsApiAddSheetRequestPayload {
        [JsonPropertyName("properties")]
        public GoogleSheetsApiSheetPropertiesPayload Properties { get; set; } = new GoogleSheetsApiSheetPropertiesPayload();
    }

    internal sealed class GoogleSheetsApiUpdateCellsRequestPayload {
        [JsonPropertyName("start")]
        public GoogleSheetsApiGridCoordinatePayload Start { get; set; } = new GoogleSheetsApiGridCoordinatePayload();

        [JsonPropertyName("rows")]
        public List<GoogleSheetsApiRowDataPayload> Rows { get; set; } = new List<GoogleSheetsApiRowDataPayload>();

        [JsonPropertyName("fields")]
        public string Fields { get; set; } = "userEnteredValue";
    }

    internal sealed class GoogleSheetsApiSetDataValidationRequestPayload {
        [JsonPropertyName("range")]
        public GoogleSheetsApiGridRangePayload Range { get; set; } = new GoogleSheetsApiGridRangePayload();

        [JsonPropertyName("rule")]
        public GoogleSheetsApiDataValidationRulePayload Rule { get; set; } = new GoogleSheetsApiDataValidationRulePayload();
    }

    internal sealed class GoogleSheetsApiGridCoordinatePayload {
        [JsonPropertyName("sheetId")]
        public int SheetId { get; set; }

        [JsonPropertyName("rowIndex")]
        public int RowIndex { get; set; }

        [JsonPropertyName("columnIndex")]
        public int ColumnIndex { get; set; }
    }

    internal sealed class GoogleSheetsApiRowDataPayload {
        [JsonPropertyName("values")]
        public List<GoogleSheetsApiCellDataPayload> Values { get; } = new List<GoogleSheetsApiCellDataPayload>();
    }

    internal sealed class GoogleSheetsApiCellDataPayload {
        [JsonPropertyName("userEnteredValue")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiExtendedValuePayload? UserEnteredValue { get; set; }

        [JsonPropertyName("userEnteredFormat")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiCellFormatPayload? UserEnteredFormat { get; set; }

        [JsonPropertyName("dataValidationRule")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiDataValidationRulePayload? DataValidationRule { get; set; }

        [JsonPropertyName("note")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Note { get; set; }

        [JsonPropertyName("textFormatRuns")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public List<GoogleSheetsApiTextFormatRunPayload>? TextFormatRuns { get; set; }

        [JsonPropertyName("pivotTable")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiPivotTablePayload? PivotTable { get; set; }
    }

    internal sealed class GoogleSheetsApiExtendedValuePayload {
        [JsonPropertyName("stringValue")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? StringValue { get; set; }

        [JsonPropertyName("numberValue")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public double? NumberValue { get; set; }

        [JsonPropertyName("boolValue")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? BoolValue { get; set; }

        [JsonPropertyName("formulaValue")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? FormulaValue { get; set; }
    }

    internal sealed class GoogleSheetsApiCellFormatPayload {
        [JsonPropertyName("numberFormat")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiNumberFormatPayload? NumberFormat { get; set; }

        [JsonPropertyName("backgroundColor")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiColorPayload? BackgroundColor { get; set; }

        [JsonPropertyName("borders")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiBordersPayload? Borders { get; set; }

        [JsonPropertyName("textFormat")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiTextFormatPayload? TextFormat { get; set; }

        [JsonPropertyName("horizontalAlignment")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? HorizontalAlignment { get; set; }

        [JsonPropertyName("verticalAlignment")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? VerticalAlignment { get; set; }

        [JsonPropertyName("wrapStrategy")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? WrapStrategy { get; set; }

        [JsonPropertyName("padding")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiPaddingPayload? Padding { get; set; }

        [JsonPropertyName("textRotation")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiTextRotationPayload? TextRotation { get; set; }
    }

    internal sealed class GoogleSheetsApiNumberFormatPayload {
        [JsonPropertyName("type")]
        public string Type { get; set; } = "NUMBER";

        [JsonPropertyName("pattern")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Pattern { get; set; }
    }

    internal sealed class GoogleSheetsApiTextFormatPayload {
        [JsonPropertyName("bold")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? Bold { get; set; }

        [JsonPropertyName("italic")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? Italic { get; set; }

        [JsonPropertyName("underline")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? Underline { get; set; }

        [JsonPropertyName("strikethrough")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? Strikethrough { get; set; }

        [JsonPropertyName("fontFamily")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? FontFamily { get; set; }

        [JsonPropertyName("fontSize")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? FontSize { get; set; }

        [JsonPropertyName("foregroundColor")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiColorPayload? ForegroundColor { get; set; }
    }

    internal sealed class GoogleSheetsApiBatchUpdateValuesPayload {
        [JsonPropertyName("valueInputOption")]
        public string ValueInputOption { get; set; } = "USER_ENTERED";

        [JsonPropertyName("data")]
        public List<GoogleSheetsApiValueRangePayload> Data { get; } = new List<GoogleSheetsApiValueRangePayload>();
    }

    internal sealed class GoogleSheetsApiValueRangePayload {
        [JsonPropertyName("range")]
        public string Range { get; set; } = string.Empty;

        [JsonPropertyName("majorDimension")]
        public string MajorDimension { get; set; } = "ROWS";

        [JsonPropertyName("values")]
        public List<List<object?>> Values { get; } = new List<List<object?>>();
    }

    internal sealed class GoogleSheetsApiPaddingPayload {
        [JsonPropertyName("top")]
        public int Top { get; set; }

        [JsonPropertyName("right")]
        public int Right { get; set; }

        [JsonPropertyName("bottom")]
        public int Bottom { get; set; }

        [JsonPropertyName("left")]
        public int Left { get; set; }
    }

    internal sealed class GoogleSheetsApiTextRotationPayload {
        [JsonPropertyName("angle")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? Angle { get; set; }

        [JsonPropertyName("vertical")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? Vertical { get; set; }
    }

    internal sealed class GoogleSheetsApiBordersPayload {
        [JsonPropertyName("top")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiBorderPayload? Top { get; set; }

        [JsonPropertyName("bottom")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiBorderPayload? Bottom { get; set; }

        [JsonPropertyName("left")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiBorderPayload? Left { get; set; }

        [JsonPropertyName("right")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiBorderPayload? Right { get; set; }
    }

    internal sealed class GoogleSheetsApiBorderPayload {
        [JsonPropertyName("style")]
        public string Style { get; set; } = "SOLID";

        [JsonPropertyName("color")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiColorPayload? Color { get; set; }
    }

    internal sealed class GoogleSheetsApiColorPayload {
        [JsonPropertyName("red")]
        public double Red { get; set; }

        [JsonPropertyName("green")]
        public double Green { get; set; }

        [JsonPropertyName("blue")]
        public double Blue { get; set; }
    }

    internal sealed class GoogleSheetsApiMergeCellsRequestPayload {
        [JsonPropertyName("range")]
        public GoogleSheetsApiGridRangePayload Range { get; set; } = new GoogleSheetsApiGridRangePayload();

        [JsonPropertyName("mergeType")]
        public string MergeType { get; set; } = "MERGE_ALL";
    }

    internal sealed class GoogleSheetsApiGridRangePayload {
        [JsonPropertyName("sheetId")]
        public int SheetId { get; set; }

        [JsonPropertyName("startRowIndex")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? StartRowIndex { get; set; }

        [JsonPropertyName("endRowIndex")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? EndRowIndex { get; set; }

        [JsonPropertyName("startColumnIndex")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? StartColumnIndex { get; set; }

        [JsonPropertyName("endColumnIndex")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? EndColumnIndex { get; set; }
    }

    internal sealed class GoogleSheetsApiAddNamedRangeRequestPayload {
        [JsonPropertyName("namedRange")]
        public GoogleSheetsApiNamedRangePayload NamedRange { get; set; } = new GoogleSheetsApiNamedRangePayload();
    }

    internal sealed class GoogleSheetsApiAddProtectedRangeRequestPayload {
        [JsonPropertyName("protectedRange")]
        public GoogleSheetsApiProtectedRangePayload ProtectedRange { get; set; } = new GoogleSheetsApiProtectedRangePayload();
    }

    internal sealed class GoogleSheetsApiSetBasicFilterRequestPayload {
        [JsonPropertyName("filter")]
        public GoogleSheetsApiBasicFilterPayload Filter { get; set; } = new GoogleSheetsApiBasicFilterPayload();
    }

    internal sealed class GoogleSheetsApiAddFilterViewRequestPayload {
        [JsonPropertyName("filter")]
        public GoogleSheetsApiFilterViewPayload Filter { get; set; } = new GoogleSheetsApiFilterViewPayload();
    }

    internal sealed class GoogleSheetsApiAddTableRequestPayload {
        [JsonPropertyName("table")]
        public GoogleSheetsApiTablePayload Table { get; set; } = new GoogleSheetsApiTablePayload();
    }

    internal sealed class GoogleSheetsApiBasicFilterPayload {
        [JsonPropertyName("range")]
        public GoogleSheetsApiGridRangePayload Range { get; set; } = new GoogleSheetsApiGridRangePayload();

        [JsonPropertyName("criteria")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public Dictionary<string, GoogleSheetsApiFilterCriteriaPayload>? Criteria { get; set; }
    }

    internal sealed class GoogleSheetsApiFilterViewPayload {
        [JsonPropertyName("title")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Title { get; set; }

        [JsonPropertyName("range")]
        public GoogleSheetsApiGridRangePayload Range { get; set; } = new GoogleSheetsApiGridRangePayload();

        [JsonPropertyName("criteria")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public Dictionary<string, GoogleSheetsApiFilterCriteriaPayload>? Criteria { get; set; }
    }

    internal sealed class GoogleSheetsApiFilterCriteriaPayload {
        [JsonPropertyName("hiddenValues")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public List<string>? HiddenValues { get; set; }

        [JsonPropertyName("condition")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiBooleanConditionPayload? Condition { get; set; }
    }

    internal sealed class GoogleSheetsApiBooleanConditionPayload {
        [JsonPropertyName("type")]
        public string Type { get; set; } = string.Empty;

        [JsonPropertyName("values")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public List<GoogleSheetsApiConditionValuePayload>? Values { get; set; }
    }

    internal sealed class GoogleSheetsApiConditionValuePayload {
        [JsonPropertyName("userEnteredValue")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? UserEnteredValue { get; set; }
    }

    internal sealed class GoogleSheetsApiTablePayload {
        [JsonPropertyName("name")]
        public string Name { get; set; } = string.Empty;

        [JsonPropertyName("range")]
        public GoogleSheetsApiGridRangePayload Range { get; set; } = new GoogleSheetsApiGridRangePayload();

        [JsonPropertyName("rowsProperties")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiTableRowsPropertiesPayload? RowsProperties { get; set; }

        [JsonPropertyName("columnProperties")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public List<GoogleSheetsApiTableColumnPropertiesPayload>? ColumnProperties { get; set; }
    }

    internal sealed class GoogleSheetsApiTableRowsPropertiesPayload {
        [JsonPropertyName("headerColorStyle")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiColorStylePayload? HeaderColorStyle { get; set; }

        [JsonPropertyName("firstBandColorStyle")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiColorStylePayload? FirstBandColorStyle { get; set; }

        [JsonPropertyName("secondBandColorStyle")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiColorStylePayload? SecondBandColorStyle { get; set; }

        [JsonPropertyName("footerColorStyle")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiColorStylePayload? FooterColorStyle { get; set; }
    }

    internal sealed class GoogleSheetsApiColorStylePayload {
        [JsonPropertyName("rgbColor")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiColorPayload? RgbColor { get; set; }
    }

    internal sealed class GoogleSheetsApiTableColumnPropertiesPayload {
        [JsonPropertyName("columnIndex")]
        public int ColumnIndex { get; set; }

        [JsonPropertyName("columnName")]
        public string ColumnName { get; set; } = string.Empty;

        [JsonPropertyName("columnType")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? ColumnType { get; set; }

        [JsonPropertyName("dataValidationRule")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiDataValidationRulePayload? DataValidationRule { get; set; }
    }

    internal sealed class GoogleSheetsApiDataValidationRulePayload {
        [JsonPropertyName("condition")]
        public GoogleSheetsApiBooleanConditionPayload Condition { get; set; } = new GoogleSheetsApiBooleanConditionPayload();

        [JsonPropertyName("strict")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool Strict { get; set; }

        [JsonPropertyName("showCustomUi")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool ShowCustomUi { get; set; }
    }

    internal sealed class GoogleSheetsApiNamedRangePayload {
        [JsonPropertyName("name")]
        public string Name { get; set; } = string.Empty;

        [JsonPropertyName("range")]
        public GoogleSheetsApiGridRangePayload Range { get; set; } = new GoogleSheetsApiGridRangePayload();
    }

    internal sealed class GoogleSheetsApiProtectedRangePayload {
        [JsonPropertyName("range")]
        public GoogleSheetsApiGridRangePayload Range { get; set; } = new GoogleSheetsApiGridRangePayload();

        [JsonPropertyName("description")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Description { get; set; }

        [JsonPropertyName("warningOnly")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool WarningOnly { get; set; }

        [JsonPropertyName("editors")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleSheetsApiEditorsPayload? Editors { get; set; }

        [JsonPropertyName("unprotectedRanges")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public List<GoogleSheetsApiGridRangePayload>? UnprotectedRanges { get; set; }
    }

    internal sealed class GoogleSheetsApiUpdateDimensionPropertiesRequestPayload {
        [JsonPropertyName("range")]
        public GoogleSheetsApiDimensionRangePayload Range { get; set; } = new GoogleSheetsApiDimensionRangePayload();

        [JsonPropertyName("properties")]
        public GoogleSheetsApiDimensionPropertiesPayload Properties { get; set; } = new GoogleSheetsApiDimensionPropertiesPayload();

        [JsonPropertyName("fields")]
        public string Fields { get; set; } = "pixelSize";
    }

    internal sealed class GoogleSheetsApiDeleteSheetRequestPayload {
        [JsonPropertyName("sheetId")]
        public int SheetId { get; set; }
    }

    internal sealed class GoogleSheetsApiUpdateSpreadsheetPropertiesRequestPayload {
        [JsonPropertyName("properties")]
        public GoogleSheetsApiSpreadsheetPropertiesPayload Properties { get; set; } = new GoogleSheetsApiSpreadsheetPropertiesPayload();

        [JsonPropertyName("fields")]
        public string Fields { get; set; } = "title";
    }

    internal sealed class GoogleSheetsApiDimensionRangePayload {
        [JsonPropertyName("sheetId")]
        public int SheetId { get; set; }

        [JsonPropertyName("dimension")]
        public string Dimension { get; set; } = "ROWS";

        [JsonPropertyName("startIndex")]
        public int StartIndex { get; set; }

        [JsonPropertyName("endIndex")]
        public int EndIndex { get; set; }
    }

    internal sealed class GoogleSheetsApiDimensionPropertiesPayload {
        [JsonPropertyName("pixelSize")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? PixelSize { get; set; }

        [JsonPropertyName("hiddenByUser")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? HiddenByUser { get; set; }
    }
}
