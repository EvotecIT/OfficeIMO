namespace OfficeIMO.Studio.Features.Editor;

public enum PdfEditorTool {
    Select,
    Note,
    FreeText,
    Highlight,
    Underline,
    StrikeOut,
    Rectangle,
    Ellipse,
    Line,
    Ink,
    Stamp,
    AddText,
    AddImage,
    Link,
    SignatureAppearance,
    Redact
}

public sealed record PdfEditorToolChoice(PdfEditorTool Tool, string Label, string Hint);
