namespace OfficeIMO.Studio.Features.Editor;

internal readonly record struct PdfEditorVisualPoint(double X, double Y);

internal sealed record PdfEditorGesture(
    int PageNumber,
    double Left,
    double Top,
    double Right,
    double Bottom,
    IReadOnlyList<PdfEditorVisualPoint> Path);

internal sealed record PdfEditorSelection(int PageNumber, int ObjectNumber, string Subtype);
