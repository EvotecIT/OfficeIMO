using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Reader;

internal sealed record PdfObjectTransformGesture(PdfEditorSelection Selection, PdfEditorVisualBounds Target,
    IReadOnlyList<PdfEditorSelection>? Annotations = null);
