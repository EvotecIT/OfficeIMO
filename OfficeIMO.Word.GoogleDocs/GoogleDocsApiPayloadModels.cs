using System.Text.Json.Serialization;

namespace OfficeIMO.Word.GoogleDocs {
    internal sealed class GoogleDocsApiCreateDocumentPayload {
        [JsonPropertyName("title")]
        public string? Title { get; set; }
    }

    internal sealed class GoogleDocsApiBatchUpdatePayload {
        [JsonPropertyName("requests")]
        public List<GoogleDocsApiRequestPayload> Requests { get; } = new List<GoogleDocsApiRequestPayload>();
    }

    internal sealed class GoogleDocsApiRequestPayload {
        [JsonPropertyName("insertText")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiInsertTextRequestPayload? InsertText { get; set; }

        [JsonPropertyName("createHeader")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiCreateHeaderRequestPayload? CreateHeader { get; set; }

        [JsonPropertyName("createFooter")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiCreateFooterRequestPayload? CreateFooter { get; set; }

        [JsonPropertyName("insertInlineImage")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiInsertInlineImageRequestPayload? InsertInlineImage { get; set; }

        [JsonPropertyName("updateTextStyle")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiUpdateTextStyleRequestPayload? UpdateTextStyle { get; set; }

        [JsonPropertyName("updateParagraphStyle")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiUpdateParagraphStyleRequestPayload? UpdateParagraphStyle { get; set; }

        [JsonPropertyName("createParagraphBullets")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiCreateParagraphBulletsRequestPayload? CreateParagraphBullets { get; set; }

        [JsonPropertyName("insertTable")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiInsertTableRequestPayload? InsertTable { get; set; }

        [JsonPropertyName("mergeTableCells")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiMergeTableCellsRequestPayload? MergeTableCells { get; set; }

        [JsonPropertyName("insertPageBreak")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiInsertPageBreakRequestPayload? InsertPageBreak { get; set; }

        [JsonPropertyName("insertSectionBreak")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiInsertSectionBreakRequestPayload? InsertSectionBreak { get; set; }

        [JsonPropertyName("deleteContentRange")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiDeleteContentRangeRequestPayload? DeleteContentRange { get; set; }
    }

    internal sealed class GoogleDocsApiInsertTextRequestPayload {
        [JsonPropertyName("location")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiLocationPayload? Location { get; set; }

        [JsonPropertyName("endOfSegmentLocation")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiEndOfSegmentLocationPayload? EndOfSegmentLocation { get; set; }

        [JsonPropertyName("text")]
        public string Text { get; set; } = string.Empty;
    }

    internal sealed class GoogleDocsApiCreateHeaderRequestPayload {
        [JsonPropertyName("type")]
        public string Type { get; set; } = "DEFAULT";

        [JsonPropertyName("sectionBreakLocation")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiLocationPayload? SectionBreakLocation { get; set; }
    }

    internal sealed class GoogleDocsApiCreateFooterRequestPayload {
        [JsonPropertyName("type")]
        public string Type { get; set; } = "DEFAULT";

        [JsonPropertyName("sectionBreakLocation")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiLocationPayload? SectionBreakLocation { get; set; }
    }

    internal sealed class GoogleDocsApiInsertInlineImageRequestPayload {
        [JsonPropertyName("uri")]
        public string Uri { get; set; } = string.Empty;

        [JsonPropertyName("objectSize")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiSizePayload? ObjectSize { get; set; }

        [JsonPropertyName("location")]
        public GoogleDocsApiLocationPayload Location { get; set; } = new GoogleDocsApiLocationPayload();
    }

    internal sealed class GoogleDocsApiUpdateTextStyleRequestPayload {
        [JsonPropertyName("range")]
        public GoogleDocsApiRangePayload Range { get; set; } = new GoogleDocsApiRangePayload();

        [JsonPropertyName("textStyle")]
        public GoogleDocsApiTextStylePayload TextStyle { get; set; } = new GoogleDocsApiTextStylePayload();

        [JsonPropertyName("fields")]
        public string Fields { get; set; } = string.Empty;
    }

    internal sealed class GoogleDocsApiUpdateParagraphStyleRequestPayload {
        [JsonPropertyName("range")]
        public GoogleDocsApiRangePayload Range { get; set; } = new GoogleDocsApiRangePayload();

        [JsonPropertyName("paragraphStyle")]
        public GoogleDocsApiParagraphStylePayload ParagraphStyle { get; set; } = new GoogleDocsApiParagraphStylePayload();

        [JsonPropertyName("fields")]
        public string Fields { get; set; } = string.Empty;
    }

    internal sealed class GoogleDocsApiCreateParagraphBulletsRequestPayload {
        [JsonPropertyName("range")]
        public GoogleDocsApiRangePayload Range { get; set; } = new GoogleDocsApiRangePayload();

        [JsonPropertyName("bulletPreset")]
        public string BulletPreset { get; set; } = "BULLET_DISC_CIRCLE_SQUARE";
    }

    internal sealed class GoogleDocsApiInsertTableRequestPayload {
        [JsonPropertyName("rows")]
        public int Rows { get; set; }

        [JsonPropertyName("columns")]
        public int Columns { get; set; }

        [JsonPropertyName("location")]
        public GoogleDocsApiLocationPayload Location { get; set; } = new GoogleDocsApiLocationPayload();
    }

    internal sealed class GoogleDocsApiMergeTableCellsRequestPayload {
        [JsonPropertyName("tableRange")]
        public GoogleDocsApiTableRangePayload TableRange { get; set; } = new GoogleDocsApiTableRangePayload();
    }

    internal sealed class GoogleDocsApiInsertPageBreakRequestPayload {
        [JsonPropertyName("location")]
        public GoogleDocsApiLocationPayload Location { get; set; } = new GoogleDocsApiLocationPayload();
    }

    internal sealed class GoogleDocsApiInsertSectionBreakRequestPayload {
        [JsonPropertyName("sectionType")]
        public string SectionType { get; set; } = "NEXT_PAGE";

        [JsonPropertyName("location")]
        public GoogleDocsApiLocationPayload Location { get; set; } = new GoogleDocsApiLocationPayload();
    }

    internal sealed class GoogleDocsApiDeleteContentRangeRequestPayload {
        [JsonPropertyName("range")]
        public GoogleDocsApiRangePayload Range { get; set; } = new GoogleDocsApiRangePayload();
    }

    internal sealed class GoogleDocsApiLocationPayload {
        [JsonPropertyName("index")]
        public int Index { get; set; }

        [JsonPropertyName("segmentId")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? SegmentId { get; set; }
    }

    internal sealed class GoogleDocsApiRangePayload {
        [JsonPropertyName("startIndex")]
        public int StartIndex { get; set; }

        [JsonPropertyName("endIndex")]
        public int EndIndex { get; set; }

        [JsonPropertyName("segmentId")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? SegmentId { get; set; }
    }

    internal sealed class GoogleDocsApiEndOfSegmentLocationPayload {
        [JsonPropertyName("segmentId")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? SegmentId { get; set; }
    }

    internal sealed class GoogleDocsApiTableRangePayload {
        [JsonPropertyName("tableCellLocation")]
        public GoogleDocsApiTableCellLocationPayload TableCellLocation { get; set; } = new GoogleDocsApiTableCellLocationPayload();

        [JsonPropertyName("rowSpan")]
        public int RowSpan { get; set; }

        [JsonPropertyName("columnSpan")]
        public int ColumnSpan { get; set; }
    }

    internal sealed class GoogleDocsApiTableCellLocationPayload {
        [JsonPropertyName("tableStartLocation")]
        public GoogleDocsApiLocationPayload TableStartLocation { get; set; } = new GoogleDocsApiLocationPayload();

        [JsonPropertyName("rowIndex")]
        public int RowIndex { get; set; }

        [JsonPropertyName("columnIndex")]
        public int ColumnIndex { get; set; }
    }

    internal sealed class GoogleDocsApiTextStylePayload {
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

        [JsonPropertyName("fontSize")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiDimensionPayload? FontSize { get; set; }

        [JsonPropertyName("foregroundColor")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiOptionalColorPayload? ForegroundColor { get; set; }

        [JsonPropertyName("link")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiLinkPayload? Link { get; set; }
    }

    internal sealed class GoogleDocsApiParagraphStylePayload {
        [JsonPropertyName("namedStyleType")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? NamedStyleType { get; set; }

        [JsonPropertyName("alignment")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Alignment { get; set; }
    }

    internal sealed class GoogleDocsApiDimensionPayload {
        [JsonPropertyName("magnitude")]
        public double Magnitude { get; set; }

        [JsonPropertyName("unit")]
        public string Unit { get; set; } = "PT";
    }

    internal sealed class GoogleDocsApiSizePayload {
        [JsonPropertyName("height")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiDimensionPayload? Height { get; set; }

        [JsonPropertyName("width")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiDimensionPayload? Width { get; set; }
    }

    internal sealed class GoogleDocsApiOptionalColorPayload {
        [JsonPropertyName("color")]
        public GoogleDocsApiColorPayload Color { get; set; } = new GoogleDocsApiColorPayload();
    }

    internal sealed class GoogleDocsApiColorPayload {
        [JsonPropertyName("rgbColor")]
        public GoogleDocsApiRgbColorPayload RgbColor { get; set; } = new GoogleDocsApiRgbColorPayload();
    }

    internal sealed class GoogleDocsApiRgbColorPayload {
        [JsonPropertyName("red")]
        public double Red { get; set; }

        [JsonPropertyName("green")]
        public double Green { get; set; }

        [JsonPropertyName("blue")]
        public double Blue { get; set; }
    }

    internal sealed class GoogleDocsApiLinkPayload {
        [JsonPropertyName("url")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Url { get; set; }
    }

    internal sealed class GoogleDocsApiCreateDocumentResponse {
        [JsonPropertyName("documentId")]
        public string? DocumentId { get; set; }

        [JsonPropertyName("title")]
        public string? Title { get; set; }
    }

    internal sealed class GoogleDocsApiDocumentResponse {
        [JsonPropertyName("documentId")]
        public string? DocumentId { get; set; }

        [JsonPropertyName("title")]
        public string? Title { get; set; }

        [JsonPropertyName("body")]
        public GoogleDocsApiBodyResponse? Body { get; set; }

        [JsonPropertyName("headers")]
        public Dictionary<string, GoogleDocsApiHeaderFooterResponse>? Headers { get; set; }

        [JsonPropertyName("footers")]
        public Dictionary<string, GoogleDocsApiHeaderFooterResponse>? Footers { get; set; }
    }

    internal sealed class GoogleDocsApiBatchUpdateResponse {
        [JsonPropertyName("replies")]
        public List<GoogleDocsApiBatchUpdateReplyPayload> Replies { get; set; } = new List<GoogleDocsApiBatchUpdateReplyPayload>();
    }

    internal sealed class GoogleDocsApiBatchUpdateReplyPayload {
        [JsonPropertyName("createHeader")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiCreateHeaderResponsePayload? CreateHeader { get; set; }

        [JsonPropertyName("createFooter")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiCreateFooterResponsePayload? CreateFooter { get; set; }
    }

    internal sealed class GoogleDocsApiCreateHeaderResponsePayload {
        [JsonPropertyName("headerId")]
        public string? HeaderId { get; set; }
    }

    internal sealed class GoogleDocsApiCreateFooterResponsePayload {
        [JsonPropertyName("footerId")]
        public string? FooterId { get; set; }
    }

    internal sealed class GoogleDocsApiBodyResponse {
        [JsonPropertyName("content")]
        public List<GoogleDocsApiStructuralElementResponse>? Content { get; set; }
    }

    internal sealed class GoogleDocsApiHeaderFooterResponse {
        [JsonPropertyName("content")]
        public List<GoogleDocsApiStructuralElementResponse>? Content { get; set; }
    }

    internal sealed class GoogleDocsApiStructuralElementResponse {
        [JsonPropertyName("startIndex")]
        public int? StartIndex { get; set; }

        [JsonPropertyName("endIndex")]
        public int? EndIndex { get; set; }

        [JsonPropertyName("paragraph")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiParagraphElementResponse? Paragraph { get; set; }

        [JsonPropertyName("table")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiTableResponse? Table { get; set; }

        [JsonPropertyName("sectionBreak")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public GoogleDocsApiSectionBreakResponse? SectionBreak { get; set; }
    }

    internal sealed class GoogleDocsApiParagraphElementResponse {
    }

    internal sealed class GoogleDocsApiSectionBreakResponse {
    }

    internal sealed class GoogleDocsApiTableResponse {
        [JsonPropertyName("tableRows")]
        public List<GoogleDocsApiTableRowResponse> Rows { get; set; } = new List<GoogleDocsApiTableRowResponse>();
    }

    internal sealed class GoogleDocsApiTableRowResponse {
        [JsonPropertyName("tableCells")]
        public List<GoogleDocsApiTableCellResponse> Cells { get; set; } = new List<GoogleDocsApiTableCellResponse>();
    }

    internal sealed class GoogleDocsApiTableCellResponse {
        [JsonPropertyName("content")]
        public List<GoogleDocsApiStructuralElementResponse>? Content { get; set; }
    }
}
