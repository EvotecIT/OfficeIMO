using System.Text.Json.Serialization;

namespace OfficeIMO.Excel.GoogleSheets {
    internal sealed class GoogleSheetsApiEditorsPayload {
        [JsonPropertyName("users")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public List<string>? Users { get; set; }

        [JsonPropertyName("domainUsersCanEdit")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool DomainUsersCanEdit { get; set; }
    }

    internal sealed class GoogleSheetsApiTextFormatRunPayload {
        [JsonPropertyName("startIndex")]
        public int StartIndex { get; set; }

        [JsonPropertyName("format")]
        public GoogleSheetsApiTextFormatPayload Format { get; set; } = new GoogleSheetsApiTextFormatPayload();
    }

    internal sealed class GoogleSheetsApiAddConditionalFormatRuleRequestPayload {
        [JsonPropertyName("rule")]
        public GoogleSheetsApiConditionalFormatRulePayload Rule { get; set; } = new GoogleSheetsApiConditionalFormatRulePayload();
        [JsonPropertyName("index")]
        public int Index { get; set; }
    }

    internal sealed class GoogleSheetsApiConditionalFormatRulePayload {
        [JsonPropertyName("ranges")]
        public List<GoogleSheetsApiGridRangePayload> Ranges { get; } = new List<GoogleSheetsApiGridRangePayload>();
        [JsonPropertyName("booleanRule")]
        public GoogleSheetsApiBooleanRulePayload BooleanRule { get; set; } = new GoogleSheetsApiBooleanRulePayload();
    }

    internal sealed class GoogleSheetsApiBooleanRulePayload {
        [JsonPropertyName("condition")]
        public GoogleSheetsApiBooleanConditionPayload Condition { get; set; } = new GoogleSheetsApiBooleanConditionPayload();
        [JsonPropertyName("format")]
        public GoogleSheetsApiCellFormatPayload Format { get; set; } = new GoogleSheetsApiCellFormatPayload();
    }

    internal sealed class GoogleSheetsApiAddDimensionGroupRequestPayload {
        [JsonPropertyName("range")]
        public GoogleSheetsApiDimensionRangePayload Range { get; set; } = new GoogleSheetsApiDimensionRangePayload();
    }

    internal sealed class GoogleSheetsApiAddChartRequestPayload {
        [JsonPropertyName("chart")]
        public GoogleSheetsApiEmbeddedChartPayload Chart { get; set; } = new GoogleSheetsApiEmbeddedChartPayload();
    }

    internal sealed class GoogleSheetsApiEmbeddedChartPayload {
        [JsonPropertyName("spec")]
        public GoogleSheetsApiChartSpecPayload Spec { get; set; } = new GoogleSheetsApiChartSpecPayload();
        [JsonPropertyName("position")]
        public GoogleSheetsApiEmbeddedObjectPositionPayload Position { get; set; } = new GoogleSheetsApiEmbeddedObjectPositionPayload();
    }

    internal sealed class GoogleSheetsApiChartSpecPayload {
        [JsonPropertyName("title")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Title { get; set; }
        [JsonPropertyName("basicChart")]
        public GoogleSheetsApiBasicChartSpecPayload BasicChart { get; set; } = new GoogleSheetsApiBasicChartSpecPayload();
    }

    internal sealed class GoogleSheetsApiBasicChartSpecPayload {
        [JsonPropertyName("chartType")]
        public string ChartType { get; set; } = "COLUMN";
        [JsonPropertyName("legendPosition")]
        public string LegendPosition { get; set; } = "RIGHT_LEGEND";
        [JsonPropertyName("headerCount")]
        public int HeaderCount { get; set; } = 1;
        [JsonPropertyName("domains")]
        public List<GoogleSheetsApiBasicChartDomainPayload> Domains { get; } = new List<GoogleSheetsApiBasicChartDomainPayload>();
        [JsonPropertyName("series")]
        public List<GoogleSheetsApiBasicChartSeriesPayload> Series { get; } = new List<GoogleSheetsApiBasicChartSeriesPayload>();
    }

    internal sealed class GoogleSheetsApiBasicChartDomainPayload {
        [JsonPropertyName("domain")]
        public GoogleSheetsApiChartDataPayload Domain { get; set; } = new GoogleSheetsApiChartDataPayload();
    }

    internal sealed class GoogleSheetsApiBasicChartSeriesPayload {
        [JsonPropertyName("series")]
        public GoogleSheetsApiChartDataPayload Series { get; set; } = new GoogleSheetsApiChartDataPayload();
        [JsonPropertyName("targetAxis")]
        public string TargetAxis { get; set; } = "LEFT_AXIS";
    }

    internal sealed class GoogleSheetsApiChartDataPayload {
        [JsonPropertyName("sourceRange")]
        public GoogleSheetsApiChartSourceRangePayload SourceRange { get; set; } = new GoogleSheetsApiChartSourceRangePayload();
    }

    internal sealed class GoogleSheetsApiChartSourceRangePayload {
        [JsonPropertyName("sources")]
        public List<GoogleSheetsApiGridRangePayload> Sources { get; } = new List<GoogleSheetsApiGridRangePayload>();
    }

    internal sealed class GoogleSheetsApiEmbeddedObjectPositionPayload {
        [JsonPropertyName("overlayPosition")]
        public GoogleSheetsApiOverlayPositionPayload OverlayPosition { get; set; } = new GoogleSheetsApiOverlayPositionPayload();
    }

    internal sealed class GoogleSheetsApiOverlayPositionPayload {
        [JsonPropertyName("anchorCell")]
        public GoogleSheetsApiGridCoordinatePayload AnchorCell { get; set; } = new GoogleSheetsApiGridCoordinatePayload();
        [JsonPropertyName("widthPixels")]
        public int WidthPixels { get; set; } = 600;
        [JsonPropertyName("heightPixels")]
        public int HeightPixels { get; set; } = 371;
    }

    internal sealed class GoogleSheetsApiPivotTablePayload {
        [JsonPropertyName("source")]
        public GoogleSheetsApiGridRangePayload Source { get; set; } = new GoogleSheetsApiGridRangePayload();
        [JsonPropertyName("rows")]
        public List<GoogleSheetsApiPivotGroupPayload> Rows { get; } = new List<GoogleSheetsApiPivotGroupPayload>();
        [JsonPropertyName("columns")]
        public List<GoogleSheetsApiPivotGroupPayload> Columns { get; } = new List<GoogleSheetsApiPivotGroupPayload>();
        [JsonPropertyName("values")]
        public List<GoogleSheetsApiPivotValuePayload> Values { get; } = new List<GoogleSheetsApiPivotValuePayload>();
    }

    internal sealed class GoogleSheetsApiPivotGroupPayload {
        [JsonPropertyName("sourceColumnOffset")]
        public int SourceColumnOffset { get; set; }
        [JsonPropertyName("showTotals")]
        public bool ShowTotals { get; set; }
        [JsonPropertyName("sortOrder")]
        public string SortOrder { get; set; } = "ASCENDING";
    }

    internal sealed class GoogleSheetsApiPivotValuePayload {
        [JsonPropertyName("sourceColumnOffset")]
        public int SourceColumnOffset { get; set; }
        [JsonPropertyName("summarizeFunction")]
        public string SummarizeFunction { get; set; } = "SUM";
        [JsonPropertyName("name")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Name { get; set; }
    }

    internal sealed class GoogleSheetsApiAddDeveloperMetadataRequestPayload {
        [JsonPropertyName("developerMetadata")]
        public GoogleSheetsApiDeveloperMetadataPayload DeveloperMetadata { get; set; } = new GoogleSheetsApiDeveloperMetadataPayload();
    }

    internal sealed class GoogleSheetsApiDeveloperMetadataPayload {
        [JsonPropertyName("metadataKey")]
        public string MetadataKey { get; set; } = string.Empty;
        [JsonPropertyName("metadataValue")]
        public string MetadataValue { get; set; } = string.Empty;
        [JsonPropertyName("visibility")]
        public string Visibility { get; set; } = "DOCUMENT";
        [JsonPropertyName("location")]
        public GoogleSheetsApiDeveloperMetadataLocationPayload Location { get; set; } = new GoogleSheetsApiDeveloperMetadataLocationPayload();
    }

    internal sealed class GoogleSheetsApiDeveloperMetadataLocationPayload {
        [JsonPropertyName("spreadsheet")]
        public bool Spreadsheet { get; set; } = true;
    }
}
