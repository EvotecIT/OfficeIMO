using System.Text.Json.Serialization;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    internal sealed class GoogleSlidesApiPresentationResponse {
        [JsonPropertyName("presentationId")] public string? PresentationId { get; set; }
        [JsonPropertyName("title")] public string? Title { get; set; }
        [JsonPropertyName("revisionId")] public string? RevisionId { get; set; }
        [JsonPropertyName("pageSize")] public GoogleSlidesApiSize? PageSize { get; set; }
        [JsonPropertyName("slides")] public List<GoogleSlidesApiPage> Slides { get; set; } = new List<GoogleSlidesApiPage>();
    }

    internal sealed class GoogleSlidesApiBatchResponse {
        [JsonPropertyName("presentationId")] public string? PresentationId { get; set; }
        [JsonPropertyName("writeControl")] public GoogleSlidesApiWriteControl? WriteControl { get; set; }
    }

    internal sealed class GoogleSlidesApiWriteControl {
        [JsonPropertyName("requiredRevisionId")] public string? RequiredRevisionId { get; set; }
    }

    internal sealed class GoogleSlidesApiPage {
        [JsonPropertyName("objectId")] public string? ObjectId { get; set; }
        [JsonPropertyName("revisionId")] public string? RevisionId { get; set; }
        [JsonPropertyName("pageElements")] public List<GoogleSlidesApiPageElement> PageElements { get; set; } = new List<GoogleSlidesApiPageElement>();
        [JsonPropertyName("slideProperties")] public GoogleSlidesApiSlideProperties? SlideProperties { get; set; }
        [JsonPropertyName("notesProperties")] public GoogleSlidesApiNotesProperties? NotesProperties { get; set; }
    }

    internal sealed class GoogleSlidesApiSlideProperties {
        [JsonPropertyName("isSkipped")] public bool IsSkipped { get; set; }
        [JsonPropertyName("notesPage")] public GoogleSlidesApiPage? NotesPage { get; set; }
    }

    internal sealed class GoogleSlidesApiPageElement {
        [JsonPropertyName("objectId")] public string? ObjectId { get; set; }
        [JsonPropertyName("size")] public GoogleSlidesApiSize? Size { get; set; }
        [JsonPropertyName("transform")] public GoogleSlidesApiTransform? Transform { get; set; }
        [JsonPropertyName("shape")] public GoogleSlidesApiShape? Shape { get; set; }
        [JsonPropertyName("table")] public GoogleSlidesApiTable? Table { get; set; }
        [JsonPropertyName("image")] public GoogleSlidesApiImage? Image { get; set; }
    }

    internal sealed class GoogleSlidesApiSize {
        [JsonPropertyName("width")] public GoogleSlidesApiDimension? Width { get; set; }
        [JsonPropertyName("height")] public GoogleSlidesApiDimension? Height { get; set; }
    }

    internal sealed class GoogleSlidesApiDimension {
        [JsonPropertyName("magnitude")] public double Magnitude { get; set; }
        [JsonPropertyName("unit")] public string? Unit { get; set; }
    }

    internal sealed class GoogleSlidesApiTransform {
        [JsonPropertyName("scaleX")] public double ScaleX { get; set; } = 1;
        [JsonPropertyName("scaleY")] public double ScaleY { get; set; } = 1;
        [JsonPropertyName("translateX")] public double TranslateX { get; set; }
        [JsonPropertyName("translateY")] public double TranslateY { get; set; }
        [JsonPropertyName("unit")] public string? Unit { get; set; }
    }

    internal sealed class GoogleSlidesApiShape {
        [JsonPropertyName("shapeType")] public string? ShapeType { get; set; }
        [JsonPropertyName("text")] public GoogleSlidesApiTextContent? Text { get; set; }
    }

    internal sealed class GoogleSlidesApiTextContent {
        [JsonPropertyName("textElements")] public List<GoogleSlidesApiTextElement> TextElements { get; set; } = new List<GoogleSlidesApiTextElement>();
    }

    internal sealed class GoogleSlidesApiTextElement {
        [JsonPropertyName("textRun")] public GoogleSlidesApiTextRun? TextRun { get; set; }
    }

    internal sealed class GoogleSlidesApiTextRun {
        [JsonPropertyName("content")] public string? Content { get; set; }
        [JsonPropertyName("style")] public GoogleSlidesApiTextStyle? Style { get; set; }
    }

    internal sealed class GoogleSlidesApiTextStyle {
        [JsonPropertyName("bold")] public bool? Bold { get; set; }
        [JsonPropertyName("italic")] public bool? Italic { get; set; }
        [JsonPropertyName("underline")] public bool? Underline { get; set; }
        [JsonPropertyName("fontSize")] public GoogleSlidesApiDimension? FontSize { get; set; }
        [JsonPropertyName("fontFamily")] public string? FontFamily { get; set; }
        [JsonPropertyName("link")] public GoogleSlidesApiLink? Link { get; set; }
    }

    internal sealed class GoogleSlidesApiLink { [JsonPropertyName("url")] public string? Url { get; set; } }
    internal sealed class GoogleSlidesApiImage { [JsonPropertyName("contentUrl")] public string? ContentUrl { get; set; } [JsonPropertyName("sourceUrl")] public string? SourceUrl { get; set; } }
    internal sealed class GoogleSlidesApiTable { [JsonPropertyName("rows")] public int Rows { get; set; } [JsonPropertyName("columns")] public int Columns { get; set; } [JsonPropertyName("tableRows")] public List<GoogleSlidesApiTableRow> TableRows { get; set; } = new List<GoogleSlidesApiTableRow>(); }
    internal sealed class GoogleSlidesApiTableRow { [JsonPropertyName("tableCells")] public List<GoogleSlidesApiTableCell> TableCells { get; set; } = new List<GoogleSlidesApiTableCell>(); }
    internal sealed class GoogleSlidesApiTableCell { [JsonPropertyName("text")] public GoogleSlidesApiTextContent? Text { get; set; } }
    internal sealed class GoogleSlidesApiNotesProperties { [JsonPropertyName("speakerNotesObjectId")] public string? SpeakerNotesObjectId { get; set; } }
}
