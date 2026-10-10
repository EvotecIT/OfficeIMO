namespace OfficeIMO.Pdf;

public sealed partial class PdfPageInteractionMap {
    internal static PdfPageInteractionMap CreateAnnotationEditing(PdfPageLayoutInfo page, IReadOnlyList<PdfAnnotation> annotations) {
        var regions = annotations.Where(annotation => annotation.PageNumber == page.PageNumber &&
            annotation.Subtype is not ("Link" or "Widget" or "Popup") && annotation.ObjectNumber.HasValue &&
            annotation.HasReadableRectangle && annotation.Width > 0 && annotation.Height > 0)
            .Select(annotation => new PdfPageInteractionRegion(PdfInteractionKind.Annotation,
                page.MapUserSpaceRectangleToVisual(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2),
                subtype: annotation.Subtype, objectNumber: annotation.ObjectNumber)).ToArray();
        return new PdfPageInteractionMap(page.PageNumber, page.VisualWidth, page.VisualHeight, regions);
    }
}
