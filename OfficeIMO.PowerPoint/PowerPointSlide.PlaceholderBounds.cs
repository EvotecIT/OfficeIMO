using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    /// <summary>Resolves omitted placeholder transforms through the layout and then the master.</summary>
    internal PowerPointLayoutBox? GetExportPlaceholderBounds(PlaceholderShape placeholder) =>
        GetExportPlaceholderBounds(SlidePart.SlideLayoutPart, placeholder);

    internal static PowerPointLayoutBox? GetExportPlaceholderBounds(SlideLayoutPart? layoutPart, PlaceholderShape placeholder) {
        PlaceholderValues type = placeholder.Type?.Value ?? PlaceholderValues.Body;
        uint index = placeholder.Index?.Value ?? 0U;
        var layoutTree = layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
        OpenXmlElement? layoutShape = layoutTree?.ChildElements.FirstOrDefault(element =>
            GetLayoutPlaceholderShape(element) is { } candidate && (candidate.Index?.Value ?? 0U) == index);
        if (layoutShape != null) {
            PowerPointLayoutBox? local = GetLayoutElementBounds(layoutShape);
            if (local.HasValue) return local;
            type = GetLayoutPlaceholderShape(layoutShape)!.Type?.Value ?? PlaceholderValues.Body;
        }
        // Layout placeholders match the master's semantic type, rather than the
        // slide/layout index (which commonly differs on title/content layouts).
        var masterTree = layoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree;
        OpenXmlElement? masterShape = masterTree?.ChildElements.FirstOrDefault(element =>
            GetLayoutPlaceholderShape(element) is { } candidate &&
            SamePlaceholderType(candidate.Type?.Value ?? PlaceholderValues.Body, type));
        return masterShape == null ? null : GetLayoutElementBounds(masterShape);
    }

    private static bool SamePlaceholderType(PlaceholderValues left, PlaceholderValues right) => left == right ||
        (left == PlaceholderValues.Title || left == PlaceholderValues.CenteredTitle) &&
        (right == PlaceholderValues.Title || right == PlaceholderValues.CenteredTitle);
}
