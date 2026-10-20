using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    /// <summary>Zero-based position in the owning presentation, or -1 for a detached slide.</summary>
    public int Index {
        get {
            var presentation = _slidePart.GetParentParts().OfType<PresentationPart>().FirstOrDefault();
            var ids = presentation?.Presentation?.SlideIdList?.Elements<SlideId>();
            if (ids == null) return -1;
            int index = 0;
            foreach (var id in ids) {
                var relationship = id.RelationshipId?.Value;
                if (relationship != null && ReferenceEquals(presentation!.GetPartById(relationship), _slidePart)) return index;
                index++;
            }
            return -1;
        }
    }

    /// <summary>Inspects the current layout and owning master without modifying the package.</summary>
    /// <returns>The layout description, or null when the slide has no layout.</returns>
    public PowerPointSlideLayoutInfo? GetLayoutInfo() {
        var layout = _slidePart.SlideLayoutPart;
        if (layout == null) return null;
        var master = layout.GetParentParts().OfType<SlideMasterPart>().FirstOrDefault();
        var presentation = master?.GetParentParts().OfType<PresentationPart>().FirstOrDefault();
        int masterIndex = presentation == null ? -1 : Array.IndexOf(presentation.SlideMasterParts.ToArray(), master);
        var definition = layout.SlideLayout;
        return new PowerPointSlideLayoutInfo(masterIndex, LayoutIndex,
            definition?.CommonSlideData?.Name?.Value ?? string.Empty, definition?.Type?.Value.ToOfficeIMO(),
            master?.GetIdOfPart(layout));
    }
}
