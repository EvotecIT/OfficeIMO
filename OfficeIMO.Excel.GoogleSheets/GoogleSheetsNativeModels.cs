using System.Text.Json.Serialization;

namespace OfficeIMO.Excel.GoogleSheets {
    internal sealed class GoogleSheetsNativeSpreadsheet {
        [JsonPropertyName("spreadsheetId")]
        public string? SpreadsheetId { get; set; }

        [JsonPropertyName("spreadsheetUrl")]
        public string? SpreadsheetUrl { get; set; }

        [JsonPropertyName("properties")]
        public GoogleSheetsNativeSpreadsheetProperties? Properties { get; set; }

        [JsonPropertyName("sheets")]
        public List<GoogleSheetsNativeSheet> Sheets { get; set; } = new List<GoogleSheetsNativeSheet>();

        [JsonPropertyName("namedRanges")]
        public List<GoogleSheetsNativeNamedRange> NamedRanges { get; set; } = new List<GoogleSheetsNativeNamedRange>();
    }

    internal sealed class GoogleSheetsNativeSpreadsheetProperties {
        [JsonPropertyName("title")]
        public string? Title { get; set; }

        [JsonPropertyName("locale")]
        public string? Locale { get; set; }

        [JsonPropertyName("timeZone")]
        public string? TimeZone { get; set; }
    }

    internal sealed class GoogleSheetsNativeSheet {
        [JsonPropertyName("properties")]
        public GoogleSheetsNativeSheetProperties Properties { get; set; } = new GoogleSheetsNativeSheetProperties();

        [JsonPropertyName("data")]
        public List<GoogleSheetsNativeGridData> Data { get; set; } = new List<GoogleSheetsNativeGridData>();

        [JsonPropertyName("merges")]
        public List<GoogleSheetsNativeGridRange> Merges { get; set; } = new List<GoogleSheetsNativeGridRange>();

        [JsonPropertyName("conditionalFormats")]
        public List<object> ConditionalFormats { get; set; } = new List<object>();

        [JsonPropertyName("charts")]
        public List<object> Charts { get; set; } = new List<object>();

        [JsonPropertyName("tables")]
        public List<object> Tables { get; set; } = new List<object>();

        [JsonPropertyName("filterViews")]
        public List<object> FilterViews { get; set; } = new List<object>();

        [JsonPropertyName("basicFilter")]
        public object? BasicFilter { get; set; }

        [JsonPropertyName("rowGroups")]
        public List<GoogleSheetsNativeDimensionGroup> RowGroups { get; set; } = new List<GoogleSheetsNativeDimensionGroup>();

        [JsonPropertyName("columnGroups")]
        public List<GoogleSheetsNativeDimensionGroup> ColumnGroups { get; set; } = new List<GoogleSheetsNativeDimensionGroup>();
    }

    internal sealed class GoogleSheetsNativeSheetProperties {
        [JsonPropertyName("sheetId")]
        public int SheetId { get; set; }

        [JsonPropertyName("title")]
        public string Title { get; set; } = string.Empty;

        [JsonPropertyName("index")]
        public int Index { get; set; }

        [JsonPropertyName("hidden")]
        public bool Hidden { get; set; }

        [JsonPropertyName("rightToLeft")]
        public bool RightToLeft { get; set; }

        [JsonPropertyName("tabColor")]
        public GoogleSheetsNativeColor? TabColor { get; set; }

        [JsonPropertyName("gridProperties")]
        public GoogleSheetsNativeGridProperties? GridProperties { get; set; }
    }

    internal sealed class GoogleSheetsNativeGridProperties {
        [JsonPropertyName("frozenRowCount")]
        public int FrozenRowCount { get; set; }

        [JsonPropertyName("frozenColumnCount")]
        public int FrozenColumnCount { get; set; }

        [JsonPropertyName("hideGridlines")]
        public bool HideGridlines { get; set; }
    }

    internal sealed class GoogleSheetsNativeGridData {
        [JsonPropertyName("startRow")]
        public int StartRow { get; set; }

        [JsonPropertyName("startColumn")]
        public int StartColumn { get; set; }

        [JsonPropertyName("rowData")]
        public List<GoogleSheetsNativeRowData> RowData { get; set; } = new List<GoogleSheetsNativeRowData>();

        [JsonPropertyName("rowMetadata")]
        public List<GoogleSheetsNativeDimensionProperties> RowMetadata { get; set; } = new List<GoogleSheetsNativeDimensionProperties>();

        [JsonPropertyName("columnMetadata")]
        public List<GoogleSheetsNativeDimensionProperties> ColumnMetadata { get; set; } = new List<GoogleSheetsNativeDimensionProperties>();
    }

    internal sealed class GoogleSheetsNativeDimensionProperties {
        [JsonPropertyName("hiddenByUser")]
        public bool? HiddenByUser { get; set; }

        [JsonPropertyName("pixelSize")]
        public int? PixelSize { get; set; }
    }

    internal sealed class GoogleSheetsNativeDimensionGroup {
        [JsonPropertyName("range")]
        public GoogleSheetsNativeDimensionRange Range { get; set; } = new GoogleSheetsNativeDimensionRange();

        [JsonPropertyName("depth")]
        public int Depth { get; set; }

        [JsonPropertyName("collapsed")]
        public bool Collapsed { get; set; }
    }

    internal sealed class GoogleSheetsNativeDimensionRange {
        [JsonPropertyName("sheetId")]
        public int SheetId { get; set; }

        [JsonPropertyName("dimension")]
        public string? Dimension { get; set; }

        [JsonPropertyName("startIndex")]
        public int StartIndex { get; set; }

        [JsonPropertyName("endIndex")]
        public int EndIndex { get; set; }
    }

    internal sealed class GoogleSheetsNativeRowData {
        [JsonPropertyName("values")]
        public List<GoogleSheetsNativeCellData> Values { get; set; } = new List<GoogleSheetsNativeCellData>();
    }

    internal sealed class GoogleSheetsNativeCellData {
        [JsonPropertyName("userEnteredValue")]
        public GoogleSheetsNativeExtendedValue? UserEnteredValue { get; set; }

        [JsonPropertyName("effectiveValue")]
        public GoogleSheetsNativeExtendedValue? EffectiveValue { get; set; }

        [JsonPropertyName("formattedValue")]
        public string? FormattedValue { get; set; }

        [JsonPropertyName("userEnteredFormat")]
        public GoogleSheetsNativeCellFormat? UserEnteredFormat { get; set; }

        [JsonPropertyName("textFormatRuns")]
        public List<GoogleSheetsNativeTextFormatRun> TextFormatRuns { get; set; } = new List<GoogleSheetsNativeTextFormatRun>();

        [JsonPropertyName("note")]
        public string? Note { get; set; }

        [JsonPropertyName("dataValidation")]
        public object? DataValidation { get; set; }

        [JsonPropertyName("pivotTable")]
        public object? PivotTable { get; set; }
    }

    internal sealed class GoogleSheetsNativeTextFormatRun {
        [JsonPropertyName("startIndex")]
        public int StartIndex { get; set; }

        [JsonPropertyName("format")]
        public GoogleSheetsNativeTextFormat Format { get; set; } = new GoogleSheetsNativeTextFormat();
    }

    internal sealed class GoogleSheetsNativeExtendedValue {
        [JsonPropertyName("stringValue")]
        public string? StringValue { get; set; }

        [JsonPropertyName("numberValue")]
        public double? NumberValue { get; set; }

        [JsonPropertyName("boolValue")]
        public bool? BoolValue { get; set; }

        [JsonPropertyName("formulaValue")]
        public string? FormulaValue { get; set; }

        [JsonPropertyName("errorValue")]
        public GoogleSheetsNativeErrorValue? ErrorValue { get; set; }
    }

    internal sealed class GoogleSheetsNativeErrorValue {
        [JsonPropertyName("type")]
        public string? Type { get; set; }

        [JsonPropertyName("message")]
        public string? Message { get; set; }
    }

    internal sealed class GoogleSheetsNativeCellFormat {
        [JsonPropertyName("numberFormat")]
        public GoogleSheetsNativeNumberFormat? NumberFormat { get; set; }

        [JsonPropertyName("backgroundColor")]
        public GoogleSheetsNativeColor? BackgroundColor { get; set; }

        [JsonPropertyName("textFormat")]
        public GoogleSheetsNativeTextFormat? TextFormat { get; set; }

        [JsonPropertyName("horizontalAlignment")]
        public string? HorizontalAlignment { get; set; }

        [JsonPropertyName("verticalAlignment")]
        public string? VerticalAlignment { get; set; }

        [JsonPropertyName("wrapStrategy")]
        public string? WrapStrategy { get; set; }

        [JsonPropertyName("textRotation")]
        public GoogleSheetsNativeTextRotation? TextRotation { get; set; }
    }

    internal sealed class GoogleSheetsNativeNumberFormat {
        [JsonPropertyName("type")]
        public string? Type { get; set; }

        [JsonPropertyName("pattern")]
        public string? Pattern { get; set; }
    }

    internal sealed class GoogleSheetsNativeTextFormat {
        [JsonPropertyName("foregroundColor")]
        public GoogleSheetsNativeColor? ForegroundColor { get; set; }

        [JsonPropertyName("fontFamily")]
        public string? FontFamily { get; set; }

        [JsonPropertyName("fontSize")]
        public int? FontSize { get; set; }

        [JsonPropertyName("bold")]
        public bool? Bold { get; set; }

        [JsonPropertyName("italic")]
        public bool? Italic { get; set; }

        [JsonPropertyName("strikethrough")]
        public bool? Strikethrough { get; set; }

        [JsonPropertyName("underline")]
        public bool? Underline { get; set; }
    }

    internal sealed class GoogleSheetsNativeTextRotation {
        [JsonPropertyName("angle")]
        public int? Angle { get; set; }

        [JsonPropertyName("vertical")]
        public bool Vertical { get; set; }
    }

    internal sealed class GoogleSheetsNativeColor {
        [JsonPropertyName("red")]
        public double Red { get; set; }

        [JsonPropertyName("green")]
        public double Green { get; set; }

        [JsonPropertyName("blue")]
        public double Blue { get; set; }

        [JsonPropertyName("alpha")]
        public double? Alpha { get; set; }
    }

    internal sealed class GoogleSheetsNativeNamedRange {
        [JsonPropertyName("name")]
        public string Name { get; set; } = string.Empty;

        [JsonPropertyName("range")]
        public GoogleSheetsNativeGridRange Range { get; set; } = new GoogleSheetsNativeGridRange();
    }

    internal sealed class GoogleSheetsNativeGridRange {
        [JsonPropertyName("sheetId")]
        public int SheetId { get; set; }

        [JsonPropertyName("startRowIndex")]
        public int? StartRowIndex { get; set; }

        [JsonPropertyName("endRowIndex")]
        public int? EndRowIndex { get; set; }

        [JsonPropertyName("startColumnIndex")]
        public int? StartColumnIndex { get; set; }

        [JsonPropertyName("endColumnIndex")]
        public int? EndColumnIndex { get; set; }
    }
}
