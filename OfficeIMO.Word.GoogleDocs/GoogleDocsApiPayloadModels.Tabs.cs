using System.Text.Json.Serialization;

namespace OfficeIMO.Word.GoogleDocs {
    internal sealed class GoogleDocsApiWriteControlPayload {
        [JsonPropertyName("requiredRevisionId")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? RequiredRevisionId { get; set; }
        [JsonPropertyName("targetRevisionId")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? TargetRevisionId { get; set; }
    }

    internal sealed class GoogleDocsApiBookmarkLinkPayload {
        [JsonPropertyName("id")]
        public string? Id { get; set; }
        [JsonPropertyName("tabId")]
        public string? TabId { get; set; }
    }

    internal sealed class GoogleDocsApiHeadingLinkPayload {
        [JsonPropertyName("id")]
        public string? Id { get; set; }
        [JsonPropertyName("tabId")]
        public string? TabId { get; set; }
    }

    internal sealed class GoogleDocsApiTabResponse {
        [JsonPropertyName("tabProperties")]
        public GoogleDocsApiTabPropertiesResponse Properties { get; set; } = new GoogleDocsApiTabPropertiesResponse();
        [JsonPropertyName("documentTab")]
        public GoogleDocsApiDocumentTabResponse? DocumentTab { get; set; }
        [JsonPropertyName("childTabs")]
        public List<GoogleDocsApiTabResponse> ChildTabs { get; set; } = new List<GoogleDocsApiTabResponse>();
    }

    internal sealed class GoogleDocsApiTabPropertiesResponse {
        [JsonPropertyName("tabId")]
        public string? TabId { get; set; }
        [JsonPropertyName("title")]
        public string? Title { get; set; }
        [JsonPropertyName("index")]
        public int Index { get; set; }
        [JsonPropertyName("nestingLevel")]
        public int NestingLevel { get; set; }
        [JsonPropertyName("parentTabId")]
        public string? ParentTabId { get; set; }
    }

    internal sealed class GoogleDocsApiDocumentTabResponse {
        [JsonPropertyName("body")]
        public GoogleDocsApiBodyResponse? Body { get; set; }
        [JsonPropertyName("headers")]
        public Dictionary<string, GoogleDocsApiHeaderFooterResponse>? Headers { get; set; }
        [JsonPropertyName("footers")]
        public Dictionary<string, GoogleDocsApiHeaderFooterResponse>? Footers { get; set; }
        [JsonPropertyName("footnotes")]
        public Dictionary<string, GoogleDocsApiHeaderFooterResponse>? Footnotes { get; set; }
        [JsonPropertyName("namedRanges")]
        public Dictionary<string, GoogleDocsApiNamedRangesResponse>? NamedRanges { get; set; }
    }

    internal sealed class GoogleDocsApiNamedRangesResponse {
        [JsonPropertyName("namedRanges")]
        public List<GoogleDocsApiNamedRangeResponse> NamedRanges { get; set; } = new List<GoogleDocsApiNamedRangeResponse>();
    }

    internal sealed class GoogleDocsApiNamedRangeResponse {
        [JsonPropertyName("namedRangeId")]
        public string? NamedRangeId { get; set; }
        [JsonPropertyName("name")]
        public string? Name { get; set; }
        [JsonPropertyName("ranges")]
        public List<GoogleDocsApiRangePayload> Ranges { get; set; } = new List<GoogleDocsApiRangePayload>();
    }
}
