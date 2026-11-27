using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    /// <summary>Zero-based position in the owning presentation, or -1 for a detached slide.</summary>
    public int Index {
        get {
            // A slide removed before the SDK opened is no longer in the presentation.
            if (RemovedFromPackage) return -1;
            if (IsPackageBacked && OwnerDocument?.GetPackageSlides() is { } packageSlides) return packageSlides.IndexOf(this);
            if (OwnerDocument is not { } document) return -1;
            PowerPointPartRef presentation = document.PresentationPartRef, slide = PartRef;
            OfficeXmlNode list = presentation.Root.Child(OfficeXmlNamespaces.PresentationML, "sldIdLst");
            int index = 0;
            foreach (OfficeXmlNode entry in list.Children(OfficeXmlNamespaces.PresentationML, "sldId")) {
                string? relationship = entry.Attribute(OfficeXmlNamespaces.Relationships, "id");
                if (relationship != null && presentation.TryGetRelationship(relationship, out PowerPointRelationship target)
                    && target.TargetPart.Equals(slide)) return index;
                index++;
            }
            return -1;
        }
    }

    /// <summary>Inspects the current layout and owning master without modifying the package.</summary>
    /// <returns>The layout description, or null when the slide has no layout.</returns>
    public PowerPointSlideLayoutInfo? GetLayoutInfo() {
        PowerPointPartRef layout = PartRef.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType);
        if (layout.IsNone) return null;
        PowerPointPartRef master = layout.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType);
        int masterIndex = master.IsNone || OwnerDocument == null ? -1 : OwnerDocument.MasterPartRefs().IndexOf(master);
        OfficeXmlNode root = layout.Root;
        PowerPointSlideLayoutType? type = PowerPointSlideLayoutTypeTokens.TryParse(root.Attribute("type"), out PowerPointSlideLayoutType parsed) ? parsed : null;
        return new PowerPointSlideLayoutInfo(masterIndex, LayoutIndex,
            root.Child(OfficeXmlNamespaces.PresentationML, "cSld").Attribute("name") ?? string.Empty, type,
            master.IsNone ? null : master.RelationshipIdOf(layout, PowerPointPartRef.SlideLayoutRelationshipType));
    }
}
