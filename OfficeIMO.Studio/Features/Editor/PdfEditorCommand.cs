using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Editor;

internal sealed record PdfEditorProperties(
    string Text,
    string Author,
    PdfColor Color,
    string StampName,
    string LinkUri,
    double FontSize,
    byte[]? ImageBytes = null,
    int? LinkPageNumber = null);

internal sealed record PdfEditorCommand(
    PdfEditorTool Tool,
    int PageNumber,
    PdfPageRectangle Bounds,
    IReadOnlyList<PdfPagePoint> Path,
    PdfEditorProperties Properties,
    IReadOnlyList<IReadOnlyList<PdfPagePoint>>? Strokes = null);

internal sealed record PdfVerifiedRedactionResult(
    byte[] Bytes,
    PdfRedactionPlan Plan,
    PdfRedactionEvidenceReport Evidence,
    PdfRedactionShareableSummary Summary);
