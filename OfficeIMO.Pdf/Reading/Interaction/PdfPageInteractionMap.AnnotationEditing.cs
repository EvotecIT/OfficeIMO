namespace OfficeIMO.Pdf;

public sealed partial class PdfPageInteractionMap {
    internal static PdfPageInteractionMap CreateAnnotationEditing(PdfPageLayoutInfo page, IReadOnlyList<PdfAnnotation> annotations,
        IReadOnlyList<PdfFormField>? fields = null) {
        var regions = annotations.Where(annotation => annotation.PageNumber == page.PageNumber &&
            annotation.Subtype is not ("Link" or "Widget" or "Popup") && annotation.ObjectNumber.HasValue &&
            annotation.HasReadableRectangle && annotation.Width > 0 && annotation.Height > 0)
            .Select(annotation => new PdfPageInteractionRegion(PdfInteractionKind.Annotation,
                page.MapUserSpaceRectangleToVisual(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2),
                subtype: annotation.Subtype, objectNumber: annotation.ObjectNumber)).ToList();
        foreach (var field in fields ?? []) {
            foreach (var widget in field.Widgets.Where(widget => widget.PageNumber == page.PageNumber && !widget.IsHidden && !widget.IsNoView &&
                widget.Width != 0 && widget.Height != 0 && IsFinite(widget.X1) && IsFinite(widget.Y1) &&
                IsFinite(widget.X2) && IsFinite(widget.Y2))) {
                regions.Add(new PdfPageInteractionRegion(PdfInteractionKind.FormWidget,
                    page.MapUserSpaceRectangleToVisual(Math.Min(widget.X1, widget.X2), Math.Min(widget.Y1, widget.Y2),
                        Math.Max(widget.X1, widget.X2), Math.Max(widget.Y1, widget.Y2)),
                    fieldName: field.Name, subtype: "Widget", objectNumber: widget.ObjectNumber));
            }
        }
        return new PdfPageInteractionMap(page.PageNumber, page.VisualWidth, page.VisualHeight, regions.AsReadOnly());
    }
}
