using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Editor;

internal static class PdfEditorCommandFactory {
    internal static PdfEditorCommand Create(byte[] pdf, PdfEditorTool tool, PdfEditorGesture gesture, PdfEditorProperties properties) {
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentNullException.ThrowIfNull(gesture);
        ArgumentNullException.ThrowIfNull(properties);
        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf);
        PdfLogicalPage page = logical.Pages.FirstOrDefault(candidate => candidate.PageNumber == gesture.PageNumber)
            ?? throw new ArgumentOutOfRangeException(nameof(gesture), "The editor gesture page does not exist in the PDF.");
        PdfPageRectangle bounds = page.MapVisualRectangleToUserSpace(
            gesture.Left,
            gesture.Top,
            gesture.Right,
            gesture.Bottom);
        PdfPagePoint[] path = gesture.Path
            .Select(point => page.MapVisualPointToUserSpace(point.X, point.Y))
            .ToArray();
        return new PdfEditorCommand(tool, gesture.PageNumber, bounds, path, properties);
    }
}
