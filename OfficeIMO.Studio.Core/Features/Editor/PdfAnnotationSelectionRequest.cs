namespace OfficeIMO.Studio.Features.Editor;

/// <summary>A page-local annotation selection gesture. Additive gestures toggle whole standard groups.</summary>
internal sealed record PdfAnnotationSelectionRequest(IReadOnlyList<PdfEditorSelection> Selections, bool Additive, bool Toggle = true);
